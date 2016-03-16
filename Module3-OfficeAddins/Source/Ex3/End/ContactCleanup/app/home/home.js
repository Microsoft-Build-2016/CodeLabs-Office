(function(){
    "use strict";
    var _dlg;
    
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function(reason){
        $(document).ready(function(){
            app.initialize();

            $("#btnSignin").click(function () {
                var url = "https://localhost:8443/app/auth/auth.html";
                
                Office.context.ui.displayDialogAsync(url, { height: 40, width: 40, requireHTTPS: true }, function (result) {
                    _dlg = result.value;
                    _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, dialogMessageReceived);
                });
            });
            
            $("#btnSave").click(function() {
               app.showNotification("Note", "Save pending changes to Office 365...");
            });
        });
    };
    
    function dialogMessageReceived(result) {
        if (result && JSON.parse(result.message).status === "success") {
            $("#status").html("Connected");
            
            //close the dialog and call into Graph
            _dlg.close();
            var _token = JSON.parse(result.message);
            $.ajax({
                url: "https://graph.microsoft.com/v1.0/me/contacts",
                headers: {
                    "accept": "application/json",
                    "Authorization": "Bearer " + _token.accessToken
                },
                success: function (data) {
                    var _officeTable = new Office.TableData();
                    _officeTable.headers = ["First", "Last", "PrimaryEmail", "Company", "WorkPhone", "MobilePhone"];
                    
                    $(data.value).each(function (i, e) {
                        var _item = [
                            e.givenName,
                            e.surname,
                            (e.emailAddresses.length > 0) ? e.emailAddresses[0].address : "",
                            (e.companyName) ? e.companyName : "",
                            (e.businessPhones.length > 0) ? e.businessPhones[0] : "",
                            (e.mobilePhone) ? e.mobilePhone : ""];
                        _officeTable.rows.push(_item);
                    });

                    //toggle the buttons and set the data into Excel
                    $("#btnSignin").hide();
                    $("#btnSave").show();
                    Office.context.document.setSelectedDataAsync(_officeTable, {
                        coercionType: Office.CoercionType.Table,
                        cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }]
                    }, function (asyncResult) {
                        if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
                            //create a table binding
                            Office.context.document.bindings.addFromSelectionAsync(
                                Office.BindingType.Table, { id: "ContactsBinding" },
                                function (result) {
                                    if (result.status !== Office.AsyncResultStatus.Succeeded)
                                        app.showNotification("Error", "Binding to the table");
                                    else {
                                        //get the binding
                                        Office.context.document.bindings.getByIdAsync("ContactsBinding", function(result) {
					                       if (result.status === Office.AsyncResultStatus.Succeeded) {
                                               //add BindingDataChanged handler
                                               var _binding = result.value;
                    	                       _binding.addHandlerAsync(Office.EventType.BindingDataChanged, function () { 
                                                   app.showNotification("Note", "Data changed...");
                                                   $("#btnSave").prop("disabled", false);
                                               });
                                           }
                                        });
                                    }
                                });
                        }
                        else {
                            app.showNotification("Error", "Failed to write data to Excel");
                        }
                    });
                },
                error: function (err) {
                    app.showNotification("Error", "Failed calling Microsoft Graph");
                }
            });
        }
        else {
            app.showNotification("Error", "Failed to get token");
        }
    };
})();