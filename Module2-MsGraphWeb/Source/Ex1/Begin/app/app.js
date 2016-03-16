angular.module("myContacts.services", [])
.factory("o365Service", ["$http", "$q", function($http, $q) {
    var o365Service = {};
    
    o365Service.getContacts = function() {
        var deferred = $q.defer();
        
        var data = { value: [
            { displayName: "Richard diZerega", emailAddresses: [{ address: "ridize@rzdemos.com" }] },
            { displayName: "Andrew Coates", emailAddresses: [{ address: "acoates@rzdemos.com" }] },
            { displayName: "Jeremy Thake", emailAddresses: [{ address: "jthake@rzdemos.com" }] },
            { displayName: "Rob Howard", emailAddresses: [{ address: "rhoward@rzdemos.com" }] }
        ]};
        deferred.resolve(data);
        
        return deferred.promise;
    }
    
    return o365Service;
}]);

angular.module("myContacts.controllers", [])
.controller("loginCtrl", ["$scope", "$location", function($scope, $location) {
    $scope.login = function() {
        $location.path("/contacts");
    };
}])
.controller("contactsCtrl", ["$scope", "o365Service", function($scope, o365Service) {
    o365Service.getContacts().then(function(data) {
       $scope.contacts = data.value; 
    }, function(err) {
        $scope.error = err;
    });
    
    $scope.dismiss = function() {
      $scope.error = null;  
    };
}]);

angular.module("myContacts", ["myContacts.services", "myContacts.controllers", "ngRoute"])
.config(["$routeProvider", "$httpProvider", function($routeProvider, $httpProvider) {
    $routeProvider.when("/login", {
        controller: "loginCtrl",
        templateUrl: "/app/templates/view-login.html"
    }).when ("/contacts", {
        controller: "contactsCtrl",
        templateUrl: "/app/templates/view-contacts.html"
    }).otherwise({
        redirectTo: "/login"
    });
}]);