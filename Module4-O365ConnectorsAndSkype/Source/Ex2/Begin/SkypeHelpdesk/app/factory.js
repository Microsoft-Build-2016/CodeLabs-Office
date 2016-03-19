angular.module("helpdesk.services", [])
.factory("helpdeskSvc", ["$rootScope", "$http", "$q", function($rootScope, $http, $q) {
    var helpdeskSvc = {};
    
    //gets helpdesk tickets that are open
    helpdeskSvc.getTickets = function() {
        var deferred = $q.defer();
        
        deferred.resolve(tickets);
        
        return deferred.promise;  
    };
    
    //gets the users profile picture from Microsoft Graph
    helpdeskSvc.getProfilePic = function() {
        var deferred = $q.defer();
        
        $http.get("https://graph.microsoft.com/v1.0/me/photo/$value", { responseType: "blob" }).then(function (image) {
            // Convert blob into image that app can display
            var imgUrl = window.URL || window.webkitURL;
            var blobUrl = imgUrl.createObjectURL(image.data);
            deferred.resolve(blobUrl);
        }, function (err) {
            // Error calling API...reject the promise
            deferred.reject("Image failed to load");
        });
        
        return deferred.promise;
    };
    
    //Hack...will use hard-coded tickets for demo purposes
    var tenantDomain = "TENANT.onmicrosoft.com"; //TODO: CHANGE THIS TO YOUR OFFICE 365 TENANT DOMAIN
    var skypeTestUser = "USER2@TENANT.onmicrosoft.com"; //TODO: CHANGE THIS TO THE USER THAT WILL TEST FROM SKYPE
    var tickets = [
        { 
            id: 1,
            title: "Wifi not working in Building 35", 
            category: "Networking", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "TEST ACCOUNT", 
                email: skypeTestUser,
                status: "Offline" 
            },
            assigned_to: {}
        },
        { 
            id: 2,
            title: "Laptop screen cracked", 
            category: "Hardware", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Katie Jordon", 
                email: "katiej@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 3,
            title: "Lost phone in taxi", 
            category: "Phone", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Sara Davis", 
                email: "sarad@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 4,
            title: "Need VPN access", 
            category: "Telephony", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Rob Young", 
                email: "roby@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 5,
            title: "Need access to building 25", 
            category: "Facilities", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Rob Young", 
                email: "roby@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 6,
            title: "Software won't install", 
            category: "Software", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Garth Fort", 
                email: "garthf@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 7,
            title: "Need a landline for my new office", 
            category: "Telephony", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "TEST ACCOUNT", 
                email: skypeTestUser,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 8,
            title: "Poor Wifi in Building 10", 
            category: "Networking", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Katie Jordon", 
                email: "katiej@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 9,
            title: "Virus detected on laptop", 
            category: "Hardware", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Brian Johnson", 
                email: "brianj@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 10,
            title: "Laptop stuck on OS upgrade", 
            category: "Software", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Pavel Bansky", 
                email: "pavelb@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        },
        { 
            id: 11,
            title: "Office 365 subscription expired", 
            category: "Software", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Robin Counts", 
                email: "robinc@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        }
        ,
        { 
            id: 12,
            title: "Monitor smoking", 
            category: "Hardware", 
            status: "Open",
            created: "03-28-2016", 
            created_by: { 
                name: "Garret Vargas", 
                email: "garretv@" + tenantDomain,
                status: "Offline"
            },
            assigned_to: {}
        }
    ];
    
    return helpdeskSvc;
}]);