import Q = require("q");
import Controls = require("VSS/Controls");
import {Combo, IComboOptions} from "VSS/Controls/Combos";
import Menus = require("VSS/Controls/Menus");
import WIT_Client = require("TFS/WorkItemTracking/RestClient");
import Contracts = require("TFS/WorkItemTracking/Contracts");
import Utils_string = require("VSS/Utils/String");

import { StoredFieldReferences } from "./wsjfModels";

export class Settings {
    private _changeMade = false;
    private _selectedFields:StoredFieldReferences;
    private _fields:Contracts.WorkItemField[];
    private _menuBar = null;

    private getSortedFieldsList():IPromise<any> {
        var deferred = Q.defer();
        var client = WIT_Client.getClient();
        client.getFields().then((fields: Contracts.WorkItemField[]) => {
            this._fields = fields.filter(field => (field.type === Contracts.FieldType.Integer || field.type === Contracts.FieldType.String))
            var sortedFields = this._fields.map(field => field.name).sort((field1,field2) => {
                if (field1 > field2) {
                    return 1;
                }

                if (field1 < field2) {
                    return -1;
                }

                return 0;
            });
            deferred.resolve(sortedFields);
        });
        return deferred.promise;
    }

    private getFieldReferenceName(fieldName): string {
        let matchingFields = this._fields.filter(field => field.name === fieldName);
        return (matchingFields.length > 0) ? matchingFields[0].referenceName : null;
    }

    private getFieldName(fieldReferenceName): string {
        let matchingFields = this._fields.filter(field => field.referenceName === fieldReferenceName);
        return (matchingFields.length > 0) ? matchingFields[0].name : null;
    }

    private getComboOptions(id, fieldsList, initialField):IComboOptions {
        var that = this;
        return {
            id: id,
            mode: "drop",
            source: fieldsList,
            enabled: true,
            value: that.getFieldName(initialField),
            change: function () {
                that._changeMade = true;
                let fieldName = this.getText();
                let fieldReferenceName: string = (this.getSelectedIndex() < 0) ? null : that.getFieldReferenceName(fieldName);
                switch (this._id) {
                    case "priority":
                        that._selectedFields.priorityField = fieldReferenceName;
                        break;
                    case "risk":
                        that._selectedFields.riskField = fieldReferenceName;
                        break;
                    case "result":
                        that._selectedFields.result = fieldReferenceName;
                        break;
                }
                that.updateSaveButton();
            }
        };
    }

    public initialize() {
        let hubContent = $(".hub-content");
        let uri = VSS.getWebContext().collection.uri + "_admin/_process";

        let descriptionText = "{0} is a concept of {1} used for weighing the cost of delay with job size.";
        let header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header.html(Utils_string.format(descriptionText));

        $("<img src='https://camo.githubusercontent.com/5c45986f3b995c24393b5d2d7a8bc8038086f7e4/687474703a2f2f7777772e7363616c65646167696c656672616d65776f726b2e636f6d2f77702d636f6e74656e742f75706c6f6164732f323031342f30372f4669677572652d322e2d412d666f726d756c612d666f722d63616c63756c6174696e672d57534a462e706e67' />").addClass("description-image").appendTo(hubContent);
        
        descriptionText = "You must add a custom decimal field from the {0} to each work item type you wish to compute WSJF.";
        header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header.html(Utils_string.format(descriptionText, "<a target='_blank' href='" + uri +"'>process hub</a>"));

        let container = $("<div />").addClass("wsjf-settings-container").appendTo(hubContent);

        var menubarOptions = {
            items: [
                { id: "save", icon: "icon-save", title: "Save the selected field" }   
            ],
            executeAction:(args) => {
                var command = args.get_commandName();
                switch (command) {
                    case "save":
                        this.save();
                        break;
                    default:
                        console.log("Unhandled action: " + command);
                        break;
                }
            }
        };
        this._menuBar = Controls.create<Menus.MenuBar, any>(Menus.MenuBar, container, menubarOptions);

        let bvContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Priority").appendTo(bvContainer);

        let tcContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Risk").appendTo(tcContainer);

        let rvContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Result").appendTo(rvContainer);

        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
            dataService.getValue<StoredFieldReferences>("storedFields").then((storedFields:StoredFieldReferences) => {
                if (storedFields) {
                    console.log("Retrieved fields from storage");
                    this._selectedFields = storedFields;
                }
                else {
                    console.log("Failed to retrieve fields from storage, defaulting values")
					//Enter in your config referenceName for "rvField" and "wsjfField"
                    this._selectedFields = {
                        priorityField: "Microsoft.VSTS.Common.Priority",
                        riskField: "Microsoft.VSTS.Common.Risk",
                        result: null
                    };
                }

                this.getSortedFieldsList().then((fieldList) => {
                    Controls.create(Combo, bvContainer, this.getComboOptions("priority", fieldList, this._selectedFields.priorityField));
                    Controls.create(Combo, tcContainer, this.getComboOptions("risk", fieldList, this._selectedFields.riskField));
                    Controls.create(Combo, rvContainer, this.getComboOptions("result", fieldList, this._selectedFields.result));
                    this.updateSaveButton();
                    VSS.notifyLoadSucceeded();
                });
            });
        });
    }

    private save() {
        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
            dataService.setValue<StoredFieldReferences>("storedFields", this._selectedFields).then((storedFields:StoredFieldReferences) => {
                console.log("Storing fields completed");
                this._changeMade = false;
                this.updateSaveButton();
            });
        });
    } 

    private updateSaveButton() {
        var buttonState = (this._selectedFields.priorityField && this._selectedFields.riskField && this._selectedFields.result) && this._changeMade
                            ? Menus.MenuItemState.None : Menus.MenuItemState.Disabled;

        // Update the disabled state
        this._menuBar.updateCommandStates([
            { id: "save", disabled: buttonState },
        ]);
    }
}