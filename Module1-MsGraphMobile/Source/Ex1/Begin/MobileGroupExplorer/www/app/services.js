(function () {
    "use strict";

    angular.module("myapp.services", []).factory("myappService", ["$rootScope", "$q", function ($rootScope, $q) {
        var myappService = {};

        // Broadcasts waiting indicator toggles to app root
        myappService.wait = function (show) {
            $rootScope.$broadcast("wait", show);
        };

        // Broadcasts error messages to app root
        myappService.broadcastError = function (err) {
            $rootScope.$broadcast("error", err);
        };

        var groups = [
            { displayName: "Group 1", EmailAddress: "group1@tenant.onmicrosoft.com", img: "images/group.png" },
            { displayName: "Group 2", EmailAddress: "group2@tenant.onmicrosoft.com", img: "images/group.png" },
            { displayName: "Group 3", EmailAddress: "group3@tenant.onmicrosoft.com", img: "images/group.png" },
            { displayName: "Group 4", EmailAddress: "group4@tenant.onmicrosoft.com", img: "images/group.png" }
        ];

        myappService.getGroups = function () {
            var deferred = $q.defer();

            deferred.resolve(groups);

            return deferred.promise;
        };

        return myappService;
    }]);
})();