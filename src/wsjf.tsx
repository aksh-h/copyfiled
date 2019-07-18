import Q = require("q");

import TFS_Wit_Contracts = require("TFS/WorkItemTracking/Contracts");
import TFS_Wit_Client = require("TFS/WorkItemTracking/RestClient");
import TFS_Wit_Services = require("TFS/WorkItemTracking/Services");

import { StoredFieldReferences } from "./wsjfModels";

function GetStoredFields(): Promise<any> {
    var deferred = Q.defer();
    VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
        dataService.getValue<StoredFieldReferences>("storedFields").then((storedFields:StoredFieldReferences) => {
            if (storedFields) {
                console.log("Retrieved fields from storage");
                deferred.resolve(storedFields);
            } else {
                deferred.reject("Failed to retrieve fields from storage");
            }
        });
    });
    return deferred.promise;
}

function getWorkItemFormService()
{
    return TFS_Wit_Services.WorkItemFormService.getService();
}

function updateWSJFOnForm(storedFields:StoredFieldReferences) {
    getWorkItemFormService().then((service) => {
        service.getFields().then((fields: TFS_Wit_Contracts.WorkItemField[]) => {
            var matchingBusinessValueFields = fields.filter(field => field.referenceName === storedFields.priorityField);
            var matchingTimeCriticalityFields = fields.filter(field => field.referenceName === storedFields.riskField);
            var matchingRROEValueFields = fields.filter(field => field.referenceName === storedFields.result);

            //If this work item type has WSJF, then update WSJF
            if ((matchingBusinessValueFields.length > 0) &&
                (matchingTimeCriticalityFields.length > 0) &&
                (matchingRROEValueFields.length > 0)) {
                service.getFieldValues([storedFields.priorityField, storedFields.riskField]).then((values) => {
                    var priorityValue  = values[storedFields.priorityField];
                    var riskValue = values[storedFields.riskField];
                    var result = values[storedFields.result];

                    if (priorityValue == 1 && riskValue === "1 - High") {
                        result = "High";
                    } else if (priorityValue == 1 && riskValue === "2 - Medium") {
                        result = "Moderate";
                    } else if (priorityValue == 1 && riskValue === "3 - Low") {
                        result = "Acceptable";
                    }
                    service.setFieldValue(storedFields.result, result);
                });
            }
        });
    });
}

function updateWSJFOnGrid(workItemId, storedFields:StoredFieldReferences):IPromise<any> {
    let wsjfFields = [
        storedFields.priorityField,
        storedFields.riskField,
        storedFields.result
    ];
    var deferred = Q.defer();

    var client = TFS_Wit_Client.getClient();
    client.getWorkItem(workItemId, wsjfFields).then((workItem: TFS_Wit_Contracts.WorkItem) => {
        if (storedFields.priorityField !== undefined && storedFields.riskField !== undefined) {
            var priorityValue = workItem.fields[storedFields.priorityField];
            var riskValue = workItem.fields[storedFields.riskField];
            var result = workItem.fields[storedFields.result];

            if (priorityValue === 1 && riskValue === "1 - High") {
                result = "High";
            } else if (priorityValue === 1 && riskValue === "2 - Medium") {
                result = "Moderate";
            } else if (priorityValue === 1 && riskValue === "3 - Low") {
                result = "Acceptable";
            }

            var document = [{
                from: null,
                op: "add",
                path: '/fields/' + storedFields.result,
                value: result
            }];

            // Only update the work item if the WSJF has changed
            if (result != workItem.fields[storedFields.result]) {
                client.updateWorkItem(document, workItemId).then((updatedWorkItem:TFS_Wit_Contracts.WorkItem) => {
                    deferred.resolve(updatedWorkItem);
                });
            }
            else {
                deferred.reject("No relevant change to work item");
            }
        }
        else
        {
            deferred.reject("Unable to calculate WSJF, please configure fields on the collection settings page.");
        }
    });

    return deferred.promise;
}

var formObserver = (context) => {
    return {
        onFieldChanged: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.priorityField && storedFields.riskField && storedFields.result) {
                    //If one of fields in the calculation changes
                    if ((args.changedFields[storedFields.priorityField] !== undefined) || 
                        (args.changedFields[storedFields.riskField] !== undefined)) {
                            updateWSJFOnForm(storedFields);
                        }
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");
                }
            }, (reason) => {
                console.log(reason);
            });
        },
        
        onLoaded: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.priorityField && storedFields.riskField && storedFields.result) {
                    updateWSJFOnForm(storedFields);
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");
                }
            }, (reason) => {
                console.log(reason);
            });
        }
    } 
}

var contextProvider = (context) => {
    return {
        execute: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.priorityField && storedFields.riskField && storedFields.result) {
                    var workItemIds = args.workItemIds;
                    var promises = [];
                    $.each(workItemIds, function(index, workItemId) {
                        promises.push(updateWSJFOnGrid(workItemId, storedFields));
                    });

                    // Refresh view
                    Q.all(promises).then(() => {
                        VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: IHostNavigationService) => {
                            navigationService.reload();
                        });
                    });
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");
                    //TODO: Disable context menu item
                }
            }, (reason) => {
                console.log(reason);
            });
        }
    };
}

let extensionContext = VSS.getExtensionContext();
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.wsjf-work-item-form-observer`, formObserver);
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.wsjf-contextMenu`, contextProvider);