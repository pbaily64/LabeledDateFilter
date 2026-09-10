"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import IVisualHost               = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IFilter                   = powerbi.IFilter;

import "../style/visual.less";

// ─────────────────────────────────────────────────────────────────────────────
// Interfaces du filtre Power BI
// ─────────────────────────────────────────────────────────────────────────────
interface AdvancedFilterCondition {
    operator: string;
    value:    string;
}

interface AdvancedFilter {
    $schema:         string;
    target:          { table: string; column: string };
    logicalOperator: string;
    conditions:      AdvancedFilterCondition[];
    filterType:      number;
}

// État interne d'un champ de saisie
interface FieldState {
    input:    HTMLInputElement;
    date:     Date | null;       // null = pas de date valide saisie
    rawText:  string;            // ce que l'utilisateur a tapé (sans séparateurs)
    isValid:  boolean;           // true si date parsée correctement
    isOrderError: boolean;       // true si date2 < date1 (uniquement pour date2)
}

// ─────────────────────────────────────────────────────────────────────────────
// Visuel principal
// ─────────────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {

    private host:                      IVisualHost;
    private events: IVisualEventService;
    private container:                 HTMLElement;
    private formattingSettings:        VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private targetTable:  string = "";
    private targetColumn: string = "";

    // Inputs de saisie (1 ou 2 selon réglage)
    private field1: FieldState | null = null;
    private field2: FieldState | null = null;

    private styleEl: HTMLStyleElement | null = null;
    private dateFormat: string = "dd/mm/yyyy";

    // Empêche la restauration depuis les filtres pendant le cycle update()
    // qui suit immédiatement un applyJsonFilter émis par le visuel lui-même.
    private suppressRestore: boolean = false;

    constructor(options: VisualConstructorOptions) {
        this.host                       = options.host;
        this.events                       = options.host.eventService;
        this.formattingSettingsService  = new FormattingSettingsService();
        this.container                  = options.element as HTMLElement;
        this.container.style.overflow   = "hidden";
        this.container.style.fontFamily = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Update — appelé par Power BI à chaque changement de données ou format
    // ─────────────────────────────────────────────────────────────────────────
    // ── Rendering Events API ──────────────────────────────────────────────
    // Exigence de certification Microsoft : le visuel doit signaler a Power BI
    // le debut, la fin ou l'echec de son rendu. Sans cela, l'export et les
    // captures d'ecran peuvent intervenir avant que le rendu soit termine.
    // Le corps historique de update() est conserve tel quel dans updateCore() :
    // ses retours anticipes sortent du coeur, et renderingFinished est malgre
    // tout emis ci-dessous.
    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            this.updateCore(options);
            this.events.renderingFinished(options);
        } catch (e) {
            this.events.renderingFailed(options, e instanceof Error ? e.message : String(e));
        }
    }

    private updateCore(options: VisualUpdateOptions): void {

        const dataViews = options.dataViews;

        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, dataViews && dataViews[0]);

        const s          = this.formattingSettings;
        const showTwo    = s.dateStyle.showTwoDatePickers.value as boolean;
        const horizontal = s.dateStyle.horizontalLayout.value  as boolean;

        let label1 = (s.labelStyle.defaultLabel1.value as string) || "Date :";
        let label2 = (s.labelStyle.defaultLabel2.value as string) || "À :";

        this.dateFormat = (s.dateStyle.dateFormat.value as string) || "dd/mm/yyyy";

        // État JS courant — sert de base si aucune restauration n'a lieu
        let prevDate1 = this.field1?.date ?? null;
        let prevDate2 = this.field2?.date ?? null;

        // Lecture des labels DAX
        if (dataViews && dataViews[0]) {
            const dv = dataViews[0];

            if (dv.categorical && dv.categorical.values) {
                for (const col of dv.categorical.values) {
                    const roles = col.source && col.source.roles as Record<string, boolean>;
                    if (!roles || !col.values) continue;

                    // Cherche la première valeur non vide
                    let firstNonEmpty = "";
                    for (const v of col.values) {
                        if (v !== null && v !== undefined) {
                            const str = String(v).trim();
                            if (str !== "") { firstNonEmpty = str; break; }
                        }
                    }

                    if (roles["label1"] && firstNonEmpty !== "") { label1 = firstNonEmpty; }
                    if (roles["label2"] && firstNonEmpty !== "") { label2 = firstNonEmpty; }
                }
            }

            // Récupération de la table/colonne cible
            const cat = dv.categorical && dv.categorical.categories && dv.categorical.categories[0];
            if (cat) {
                const qn     = cat.source.queryName || "";
                const dotIdx = qn.indexOf(".");
                this.targetTable  = dotIdx >= 0 ? qn.substring(0, dotIdx) : qn;
                this.targetColumn = cat.source.displayName || "";
            }
        }

        // ── Restauration depuis les filtres actifs ───────────────────────────
        // Permet à un signet (bouton « Réinitialiser ») ou à un retour de page
        // de remettre les champs en cohérence avec le filtre réellement appliqué.
        if (this.suppressRestore) {
            this.suppressRestore = false;
        } else {
            const restored = this.readFilterDates(options);
            if (restored) {
                prevDate1 = restored.date1;
                prevDate2 = restored.date2;
            }
        }

        this.render(label1, label2, showTwo, horizontal, prevDate1, prevDate2);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Render — reconstruit le DOM
    // ─────────────────────────────────────────────────────────────────────────
    private render(
        label1: string, label2: string, showTwo: boolean, horizontal: boolean,
        prevDate1: Date | null, prevDate2: Date | null
    ): void {

        const s = this.formattingSettings;

        const bgColor       = this.colorValue(s.labelStyle.backgroundColor) || "#E66C37";
        const lblFontColor  = this.colorValue(s.labelStyle.fontColor)       || "#FFFFFF";
        const labelFontSize = Number(s.labelStyle.fontSize.value)           || 12;
        const borderRadius  = Number(s.labelStyle.borderRadius.value)       || 4;
        const dateFontSize  = Number(s.dateStyle.fontSize.value)            || 12;
        const dateFontColor = this.colorValue(s.dateStyle.fontColor)        || "#252423";

        // Vide le container
        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }

        // Style dynamique
        this.styleEl = document.createElement("style");
        this.styleEl.textContent = this.buildCSS(
            bgColor, lblFontColor, labelFontSize, borderRadius,
            dateFontSize, dateFontColor, horizontal
        );
        this.container.appendChild(this.styleEl);

        const root = document.createElement("div");
        root.className = "ldf-root";
        this.container.appendChild(root);

        // Bouton gomme
        const clearBtn = this.buildEraser(bgColor);
        clearBtn.addEventListener("click", () => this.clearAll());

        // Réinitialise les champs
        this.field1 = null;
        this.field2 = null;

        if (showTwo && horizontal) {
            // ── MODE HORIZONTAL ──────────────────────────────────────────────
            const wrapper = document.createElement("div");
            wrapper.className = "ldf-row-horizontal";

            const row1 = this.buildRow("ldf-row1", label1, "ldf-date1", prevDate1, true);
            wrapper.appendChild(row1);
            this.field1 = this.attachField(row1, "ldf-date1", prevDate1);

            const sep = document.createElement("div");
            sep.className = "ldf-separator";
            wrapper.appendChild(sep);

            const row2 = this.buildRow("ldf-row2", label2, "ldf-date2", prevDate2, true);
            row2.appendChild(clearBtn);
            wrapper.appendChild(row2);
            this.field2 = this.attachField(row2, "ldf-date2", prevDate2);

            root.appendChild(wrapper);

        } else if (showTwo) {
            // ── MODE VERTICAL, 2 sélecteurs ──────────────────────────────────
            const row1 = this.buildRow("ldf-row1", label1, "ldf-date1", prevDate1, false);
            root.appendChild(row1);
            this.field1 = this.attachField(row1, "ldf-date1", prevDate1);

            const row2 = this.buildRow("ldf-row2", label2, "ldf-date2", prevDate2, false);
            row2.appendChild(clearBtn);
            root.appendChild(row2);
            this.field2 = this.attachField(row2, "ldf-date2", prevDate2);

        } else {
            // ── MODE VERTICAL, 1 sélecteur ───────────────────────────────────
            const row1 = this.buildRow("ldf-row1", label1, "ldf-date1", prevDate1, false);
            row1.appendChild(clearBtn);
            root.appendChild(row1);
            this.field1 = this.attachField(row1, "ldf-date1", prevDate1);
        }

        // Attache les listeners de saisie
        this.attachInputListeners(showTwo);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Construction d'une ligne label + input
    // ─────────────────────────────────────────────────────────────────────────
    private buildRow(
        rowId: string, labelText: string, inputId: string,
        currentDate: Date | null, compact: boolean
    ): HTMLElement {

        const row = document.createElement("div");
        row.className = compact ? "ldf-row ldf-row--compact" : "ldf-row";
        row.id        = rowId;

        const lbl = document.createElement("div");
        lbl.className   = "ldf-label";
        lbl.textContent = labelText;

        const input = document.createElement("input");
        input.className    = "ldf-input";
        input.type         = "text";
        input.id           = inputId;
        input.placeholder  = this.formatPlaceholder(this.dateFormat);
        input.maxLength    = this.dateFormat.length;
        input.autocomplete = "off";
        input.spellcheck   = false;
        input.inputMode    = "numeric";

        if (currentDate) {
            input.value = this.toDisplayDate(currentDate, this.dateFormat);
        }

        row.appendChild(lbl);
        row.appendChild(input);
        return row;
    }

    // Crée et retourne l'état du champ
    private attachField(row: HTMLElement, inputId: string, currentDate: Date | null): FieldState {
        const input = row.querySelector("#" + inputId) as HTMLInputElement;
        const rawText = currentDate ? this.dateToRaw(currentDate) : "";
        return {
            input:        input,
            date:         currentDate,
            rawText:      rawText,
            // Un champ fraîchement construit est soit vide, soit porteur d'une
            // date valide : dans les deux cas ce n'est pas un état d'erreur.
            isValid:      true,
            isOrderError: false
        };
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Listeners de saisie (auto-formatage + validation)
    // ─────────────────────────────────────────────────────────────────────────
    private attachInputListeners(showTwo: boolean): void {

        const handleInput = (field: FieldState, ev: Event) => {
            const target = ev.target as HTMLInputElement;

            // Récupère la position du curseur AVANT modification
            const cursorBefore = target.selectionStart ?? 0;
            const valueBefore  = target.value;

            // Extrait uniquement les chiffres
            const digits = target.value.replace(/\D/g, "").substring(0, 8);
            field.rawText = digits;

            // Reformatte avec séparateurs selon dateFormat
            const formatted = this.formatRawDigits(digits, this.dateFormat);
            target.value = formatted;

            // Repositionne le curseur intelligemment
            const cursorAfter = this.computeNewCursor(valueBefore, formatted, cursorBefore);
            target.setSelectionRange(cursorAfter, cursorAfter);

            // Valide la date
            field.date    = this.parseRaw(digits, this.dateFormat);
            field.isValid = field.date !== null || digits === "";

            // Maj visuelle bordure rouge
            this.refreshFieldStyle(field);
        };

        const handleBlur = () => {
            // Validation croisée + application du filtre
            this.validateAndApply(showTwo);
        };

        const handleKeyDown = (ev: KeyboardEvent) => {
            if (ev.key === "Enter") {
                (ev.target as HTMLInputElement).blur();
            }
        };

        if (this.field1) {
            const f = this.field1;
            f.input.addEventListener("input",   (ev) => handleInput(f, ev));
            f.input.addEventListener("blur",    handleBlur);
            f.input.addEventListener("keydown", handleKeyDown);
        }
        if (this.field2) {
            const f = this.field2;
            f.input.addEventListener("input",   (ev) => handleInput(f, ev));
            f.input.addEventListener("blur",    handleBlur);
            f.input.addEventListener("keydown", handleKeyDown);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Validation croisée + application du filtre
    // ─────────────────────────────────────────────────────────────────────────
    private validateAndApply(showTwo: boolean): void {

        const f1 = this.field1;
        const f2 = this.field2;

        // Reset des erreurs d'ordre
        if (f1) { f1.isOrderError = false; }
        if (f2) { f2.isOrderError = false; }

        // Cas 1 date : on filtre dès qu'elle est valide
        if (!showTwo) {
            if (f1) { this.refreshFieldStyle(f1); }
            if (f1 && f1.isValid && f1.date) {
                this.applyFilter(f1.date, null);
            } else {
                this.clearFilter();
            }
            return;
        }

        // Cas 2 dates :
        //  - les deux doivent être saisies ET valides ET ordonnées (date2 >= date1)
        //  - sinon, aucun filtre
        if (f1 && f2) {

            // Si date1 ET date2 saisies mais date2 < date1 → marque l'erreur
            if (f1.isValid && f1.date && f2.isValid && f2.date) {
                if (f2.date < f1.date) {
                    f2.isOrderError = true;
                }
            }

            this.refreshFieldStyle(f1);
            this.refreshFieldStyle(f2);

            // Filtre seulement si tout est OK
            if (f1.isValid && f1.date && f2.isValid && f2.date && !f2.isOrderError) {
                this.applyFilter(f1.date, f2.date);
            } else {
                this.clearFilter();
            }
        }
    }

    // Met à jour la bordure d'un champ selon son état
    private refreshFieldStyle(field: FieldState): void {
        const isError = !field.isValid || field.isOrderError;
        if (isError) {
            field.input.classList.add("ldf-input--error");
        } else {
            field.input.classList.remove("ldf-input--error");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Application / suppression du filtre Power BI
    // ─────────────────────────────────────────────────────────────────────────
    /** Normalise un nom de table/colonne (crochets, apostrophes, casse). */
    private normalizeName(name: string): string {
        return (name || "").replace(/[[\]'"]/g, "").trim().toLowerCase();
    }

    /**
     * Relit options.jsonFilters pour retrouver le filtre Advanced actif sur
     * targetTable/targetColumn et en déduire les dates à afficher.
     *  - null                       → hôte sans jsonFilters ou cible inconnue
     *  - { date1: null, date2: null } → aucun filtre actif (champs à vider)
     */
    private readFilterDates(
        options: VisualUpdateOptions
    ): { date1: Date | null; date2: Date | null } | null {

        const filters = (options as unknown as { jsonFilters?: unknown[] }).jsonFilters;
        if (!Array.isArray(filters)) { return null; }
        if (!this.targetTable || !this.targetColumn) { return null; }

        const wantTable = this.normalizeName(this.targetTable);
        const wantCol   = this.normalizeName(this.targetColumn);

        let gte: Date | null = null;
        let lte: Date | null = null;

        for (const raw of filters) {
            const f = raw as {
                target?:     { table?: string; column?: string };
                conditions?: { operator?: string; value?: string }[];
            };
            if (!f || !f.target || !Array.isArray(f.conditions)) { continue; }
            if (this.normalizeName(f.target.table  || "") !== wantTable) { continue; }
            if (this.normalizeName(f.target.column || "") !== wantCol)   { continue; }

            for (const c of f.conditions) {
                if (!c || !c.value) { continue; }
                const d = new Date(c.value);
                if (isNaN(d.getTime())) { continue; }
                if (c.operator === "GreaterThanOrEqual") { gte = d; }
                if (c.operator === "LessThanOrEqual")    { lte = d; }
            }
            break;
        }

        if (gte && lte) { return { date1: gte,  date2: lte  }; }
        if (lte)        { return { date1: lte,  date2: null }; }
        if (gte)        { return { date1: gte,  date2: null }; }

        return { date1: null, date2: null };
    }

    private applyFilter(date1: Date | null, date2: Date | null): void {

        if (!this.targetTable || !this.targetColumn) { return; }

        const conditions: AdvancedFilterCondition[] = [];

        if (date1 && date2) {
            // Intervalle inclusif [date1, date2]
            const d1 = new Date(date1); d1.setHours(0, 0, 0, 0);
            const d2 = new Date(date2); d2.setHours(23, 59, 59, 999);
            conditions.push({ operator: "GreaterThanOrEqual", value: d1.toISOString() });
            conditions.push({ operator: "LessThanOrEqual",    value: d2.toISOString() });
        } else if (date1) {
            // 1 seule date → filtre date <= date1 (toute la journée)
            const d = new Date(date1); d.setHours(23, 59, 59, 999);
            conditions.push({ operator: "LessThanOrEqual", value: d.toISOString() });
        }

        if (conditions.length === 0) {
            this.clearFilter();
            return;
        }

        const filter: AdvancedFilter = {
            $schema:         "https://powerbi.com/product/schema#advanced",
            target:          { table: this.targetTable, column: this.targetColumn },
            logicalOperator: "And",
            conditions:      conditions,
            filterType:      0
        };

        this.suppressRestore = true;
        this.host.applyJsonFilter(
            filter as unknown as IFilter,
            "general", "filter",
            powerbi.FilterAction.merge
        );
    }

    private clearFilter(): void {
        this.suppressRestore = true;
        this.host.applyJsonFilter(
            [] as unknown as IFilter,
            "general", "filter",
            powerbi.FilterAction.remove
        );
    }

    private clearAll(): void {
        if (this.field1) {
            this.field1.input.value    = "";
            this.field1.rawText        = "";
            this.field1.date           = null;
            this.field1.isValid        = true;
            this.field1.isOrderError   = false;
            this.refreshFieldStyle(this.field1);
        }
        if (this.field2) {
            this.field2.input.value    = "";
            this.field2.rawText        = "";
            this.field2.date           = null;
            this.field2.isValid        = true;
            this.field2.isOrderError   = false;
            this.refreshFieldStyle(this.field2);
        }
        this.clearFilter();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Formatage et parsing des dates selon le format utilisateur
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Insère les séparateurs dans une chaîne de chiffres selon le format.
     * Ex: "12082025" + "dd/mm/yyyy" → "12/08/2025"
     *     "1208"     + "dd/mm/yyyy" → "12/08"
     *     "120"      + "dd/mm/yyyy" → "12/0"
     */
    private formatRawDigits(digits: string, format: string): string {
        if (!digits) return "";

        const tokens = this.parseFormatTokens(format);
        let result = "";
        let consumed = 0; // nombre de chiffres consommés

        for (const tok of tokens) {
            if (tok.isToken) {
                const remaining = digits.length - consumed;
                if (remaining <= 0) break;
                const take = Math.min(tok.length, remaining);
                result += digits.substring(consumed, consumed + take);
                consumed += take;
            } else {
                // Séparateur : on l'ajoute uniquement si le token qui le précède
                // a été entièrement rempli ET qu'il reste des chiffres à venir.
                const consumedAtEnd = this.consumedAtTokenEnd(tokens, tok);
                if (consumed >= consumedAtEnd && consumed < digits.length) {
                    result += tok.text;
                } else if (consumed >= consumedAtEnd && consumed === digits.length) {
                    // Cas limite : tous les chiffres saisis = pile la fin d'un token
                    // → on n'ajoute PAS le séparateur (laisse le curseur après le dernier chiffre)
                }
            }
        }
        return result;
    }

    /**
     * Renvoie le nombre total de chiffres attendus par tous les tokens
     * placés AVANT le séparateur donné dans la liste.
     */
    private consumedAtTokenEnd(
        tokens: {isToken: boolean; length: number; text: string}[],
        sepToken: {isToken: boolean; length: number; text: string}
    ): number {
        let total = 0;
        for (const tok of tokens) {
            if (tok === sepToken) break;
            if (tok.isToken) total += tok.length;
        }
        return total;
    }

    /**
     * Décompose le format en tokens. Ex: "dd/mm/yyyy" →
     *   [{isToken:true,length:2,text:'dd'},{isToken:false,text:'/'},
     *    {isToken:true,length:2,text:'mm'},{isToken:false,text:'/'},
     *    {isToken:true,length:4,text:'yyyy'}]
     */
    private parseFormatTokens(format: string): {isToken: boolean; length: number; text: string}[] {
        const tokens: {isToken: boolean; length: number; text: string}[] = [];
        const re = /(yyyy|yy|mm|dd)/gi;
        let lastIdx = 0;
        let match: RegExpExecArray | null;

        while ((match = re.exec(format)) !== null) {
            // Séparateur avant le token
            if (match.index > lastIdx) {
                const sep = format.substring(lastIdx, match.index);
                tokens.push({ isToken: false, length: sep.length, text: sep });
            }
            const t = match[0].toLowerCase();
            const len = t === "yyyy" ? 4 : t === "yy" ? 2 : 2;
            tokens.push({ isToken: true, length: len, text: t });
            lastIdx = match.index + match[0].length;
        }
        // Séparateur final éventuel
        if (lastIdx < format.length) {
            const sep = format.substring(lastIdx);
            tokens.push({ isToken: false, length: sep.length, text: sep });
        }
        return tokens;
    }

    /**
     * Recalcule la position du curseur après reformatage automatique.
     * Stratégie simple : place le curseur en fin de ce que l'utilisateur a tapé
     * (équivalent à toujours saisir à la fin, ce qui est le cas dominant).
     */
    private computeNewCursor(_valueBefore: string, formatted: string, _cursorBefore: number): number {
        // Place le curseur à la fin du texte formaté
        return formatted.length;
    }

    /**
     * Parse une chaîne de chiffres selon le format pour produire une Date.
     * Retourne null si invalide ou incomplet.
     */
    private parseRaw(digits: string, format: string): Date | null {
        if (!digits) return null;

        const tokens = this.parseFormatTokens(format).filter(t => t.isToken);

        let dd = -1, mm = -1, yyyy = -1;
        let i = 0;
        for (const tok of tokens) {
            if (i + tok.length > digits.length) {
                return null; // saisie incomplète
            }
            const slice = digits.substring(i, i + tok.length);
            i += tok.length;

            if (tok.text === "dd") {
                dd = parseInt(slice, 10);
            } else if (tok.text === "mm") {
                mm = parseInt(slice, 10);
            } else if (tok.text === "yyyy") {
                yyyy = parseInt(slice, 10);
            } else if (tok.text === "yy") {
                yyyy = 2000 + parseInt(slice, 10);
            }
        }

        if (i < digits.length) {
            return null; // trop de chiffres
        }

        // Validation des plages
        if (isNaN(dd) || isNaN(mm) || isNaN(yyyy))         return null;
        if (yyyy < 1900 || yyyy > 2999)                    return null;
        if (mm < 1 || mm > 12)                             return null;
        if (dd < 1 || dd > 31)                             return null;

        const d = new Date(yyyy, mm - 1, dd);
        // Vérifie que la date est valide (ex: 31/02 → JS la roule en 03/03)
        if (d.getFullYear() !== yyyy || d.getMonth() !== mm - 1 || d.getDate() !== dd) {
            return null;
        }
        return d;
    }

    /**
     * Convertit une Date en chaîne de chiffres pour la stocker en interne.
     */
    private dateToRaw(d: Date): string {
        const dd   = String(d.getDate()).padStart(2, "0");
        const mm   = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = String(d.getFullYear());
        const yy   = yyyy.slice(2);

        let result = "";
        const tokens = this.parseFormatTokens(this.dateFormat).filter(t => t.isToken);
        for (const tok of tokens) {
            if      (tok.text === "dd")   result += dd;
            else if (tok.text === "mm")   result += mm;
            else if (tok.text === "yyyy") result += yyyy;
            else if (tok.text === "yy")   result += yy;
        }
        return result;
    }

    /**
     * Formate une Date en respectant le format utilisateur (avec séparateurs).
     */
    private toDisplayDate(d: Date, format: string): string {
        const dd   = String(d.getDate()).padStart(2, "0");
        const mm   = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = String(d.getFullYear());
        const yy   = yyyy.slice(2);

        return format
            .replace(/yyyy/gi, yyyy)
            .replace(/yy/gi,   yy)
            .replace(/mm/gi,   mm)
            .replace(/dd/gi,   dd);
    }

    /**
     * Convertit le format en placeholder lisible (ex: "dd/mm/yyyy" → "jj/mm/aaaa").
     */
    private formatPlaceholder(format: string): string {
        return format
            .replace(/yyyy/gi, "aaaa")
            .replace(/yy/gi,   "aa")
            .replace(/mm/gi,   "mm")
            .replace(/dd/gi,   "jj");
    }

    private colorValue(picker: { value: { value: string } }): string {
        return (picker.value as { value: string })?.value ?? "";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // CSS injecté dynamiquement
    // ─────────────────────────────────────────────────────────────────────────
    private buildCSS(
        bgColor: string, lblFontColor: string, labelFontSize: number,
        borderRadius: number, dateFontSize: number, dateFontColor: string,
        horizontal: boolean
    ): string {
        return (
".ldf-root {" +
"    width: 100%;" +
"    height: 100%;" +
"    box-sizing: border-box;" +
"    padding: 6px 8px;" +
"    display: flex;" +
"    flex-direction: column;" +
"    gap: 6px;" +
"    justify-content: center;" +
"}" +
".ldf-row-horizontal {" +
"    display: flex;" +
"    flex-direction: row;" +
"    align-items: center;" +
"    gap: 8px;" +
"    width: 100%;" +
"}" +
".ldf-separator {" +
"    width: 1px;" +
"    height: 24px;" +
"    background: " + bgColor + "66;" +
"    flex-shrink: 0;" +
"}" +
".ldf-row {" +
"    display: flex;" +
"    align-items: center;" +
"    height: 32px;" +
(horizontal ? "    flex: 1;" : "    width: 100%;") +
"}" +
".ldf-row--compact {" +
"    flex: 1;" +
"    min-width: 0;" +
"}" +
".ldf-label {" +
"    display: flex;" +
"    align-items: center;" +
"    padding: 0 10px;" +
"    height: 100%;" +
"    background-color: " + bgColor + ";" +
"    color: " + lblFontColor + ";" +
"    font-size: " + labelFontSize + "px;" +
"    font-weight: 600;" +
"    border-radius: " + borderRadius + "px 0 0 " + borderRadius + "px;" +
"    white-space: nowrap;" +
"    min-width: 50px;" +
"    user-select: none;" +
"    letter-spacing: 0.02em;" +
"    box-sizing: border-box;" +
"    flex-shrink: 0;" +
"}" +
".ldf-input {" +
"    flex: 1;" +
"    min-width: 0;" +
"    height: 100%;" +
"    border: 1.5px solid " + bgColor + ";" +
"    border-left: none;" +
"    border-right: none;" +
"    border-radius: 0;" +
"    padding: 0 8px;" +
"    font-size: " + dateFontSize + "px;" +
"    color: " + dateFontColor + ";" +
"    background: #FAFAFA;" +
"    outline: none;" +
"    font-family: 'Segoe UI', sans-serif;" +
"    box-sizing: border-box;" +
"    transition: border-color 0.15s, background 0.15s;" +
"}" +
".ldf-input:hover { background: #F0F0F0; }" +
".ldf-input:focus { background: #FFFFFF; box-shadow: inset 0 0 0 2px " + bgColor + "33; }" +
".ldf-input--error {" +
"    border-color: #D13438 !important;" +
"    border-left: 1.5px solid #D13438 !important;" +
"    border-right: 1.5px solid #D13438 !important;" +
"    background: #FFF4F4;" +
"}" +
".ldf-input--error:focus { box-shadow: inset 0 0 0 2px #D1343833; }" +

/* Bouton gomme */
".ldf-eraser {" +
"    display: flex;" +
"    align-items: center;" +
"    justify-content: center;" +
"    height: 100%;" +
"    width: 30px;" +
"    flex-shrink: 0;" +
"    background: #FAFAFA;" +
"    border: 1.5px solid " + bgColor + ";" +
"    border-left: none;" +
"    border-radius: 0 " + borderRadius + "px " + borderRadius + "px 0;" +
"    cursor: pointer;" +
"    padding: 0;" +
"    box-sizing: border-box;" +
"    transition: background 0.15s;" +
"}" +
".ldf-eraser:hover { background: #FFF0E8; }" +
".ldf-eraser svg { display: block; }"
        );
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bouton gomme
    // ─────────────────────────────────────────────────────────────────────────
    private buildEraser(bgColor: string): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.className = "ldf-eraser";
        btn.title     = "Effacer les dates";

        const svgNS = "http://www.w3.org/2000/svg";
        const svg   = document.createElementNS(svgNS, "svg");
        svg.setAttribute("width",   "16");
        svg.setAttribute("height",  "16");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("fill",    "none");
        svg.setAttribute("stroke",  bgColor);
        svg.setAttribute("stroke-width",   "2");
        svg.setAttribute("stroke-linecap", "round");
        svg.setAttribute("stroke-linejoin","round");

        const body = document.createElementNS(svgNS, "path");
        body.setAttribute("d", "M20 20H7L3 16l10-10 7 7-3.5 3.5");
        svg.appendChild(body);

        const base = document.createElementNS(svgNS, "line");
        base.setAttribute("x1", "6"); base.setAttribute("y1", "20");
        base.setAttribute("x2", "20"); base.setAttribute("y2", "20");
        svg.appendChild(base);

        const diag = document.createElementNS(svgNS, "line");
        diag.setAttribute("x1", "13"); diag.setAttribute("y1", "7");
        diag.setAttribute("x2", "6");  diag.setAttribute("y2", "14");
        svg.appendChild(diag);

        btn.appendChild(svg);
        return btn;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
