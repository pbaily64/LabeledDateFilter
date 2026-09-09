"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ---------------------------------------------------------------------------
// Carte : Style des libellés
// ---------------------------------------------------------------------------
class LabelStyleCard extends FormattingSettingsCard {

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background color",
        value: { value: "#E66C37" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        value: { value: "#FFFFFF" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size (px)",
        value: 12
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border radius (px)",
        value: 4
    });

    defaultLabel1 = new formattingSettings.TextInput({
        name: "defaultLabel1",
        displayName: "Default label — date 1",
        placeholder: "ex : Before the",
        value: "Date :"
    });

    defaultLabel2 = new formattingSettings.TextInput({
        name: "defaultLabel2",
        displayName: "Default label — date 2",
        placeholder: "ex : to :",
        value: "À :"
    });

    name: string        = "labelStyle";
    displayName: string = "Labels style";
    slices: FormattingSettingsSlice[] = [
        this.backgroundColor,
        this.fontColor,
        this.fontSize,
        this.borderRadius,
        this.defaultLabel1,
        this.defaultLabel2
    ];
}

// ---------------------------------------------------------------------------
// Carte : Style des dates
// ---------------------------------------------------------------------------
class DateStyleCard extends FormattingSettingsCard {

    showTwoDatePickers = new formattingSettings.ToggleSwitch({
        name: "showTwoDatePickers",
        displayName: "Show to date pickers",
        value: false
    });

    // false = vertical (défaut) | true = horizontal (côte à côte)
    horizontalLayout = new formattingSettings.ToggleSwitch({
        name: "horizontalLayout",
        displayName: "Horizontal display (side by side)",
        value: false
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        value: { value: "#252423" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size (px)",
        value: 12
    });

    dateFormat = new formattingSettings.TextInput({
        name: "dateFormat",
        displayName: "Date format",
        placeholder: "ex : dd/mm/yyyy  or  mm-dd-yyyy",
        value: "dd/mm/yyyy"
    });

    name: string        = "dateStyle";
    displayName: string = "Dates style";
    slices: FormattingSettingsSlice[] = [
        this.showTwoDatePickers,
        this.horizontalLayout,
        this.fontColor,
        this.fontSize,
        this.dateFormat
    ];
}

// ---------------------------------------------------------------------------
// Modèle global
// ---------------------------------------------------------------------------
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    labelStyle = new LabelStyleCard();
    dateStyle  = new DateStyleCard();
    cards      = [this.labelStyle, this.dateStyle];
}
