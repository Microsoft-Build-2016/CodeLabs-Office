angular.module("myContacts.services", [])
.factory("o365Service", ["$http", "$q", function($http, $q) {
    var o365Service = {};
    
    o365Service.getContacts = function() {
        var deferred = $q.defer();
        
        $http.get("https://graph.microsoft.com/v1.0/me/contacts").then(function(result) {
            deferred.resolve(result.data);
        }, function(err) {
            deferred.reject(err.statusText);
        });
        
        return deferred.promise;
    }
    
    return o365Service;
}]);

angular.module("myContacts.controllers", [])
.controller("loginCtrl", ["$scope", "$location", "adalAuthenticationService", function($scope, $location, adalSvc) {
    if (adalSvc.userInfo.isAuthenticated) {
        $location.path("/contacts");
    }
        
    $scope.login = function() {
        adalSvc.login();  
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

angular.module("myContacts", ["myContacts.services", "myContacts.controllers", "ngRoute", "AdalAngular"])
.config(["$routeProvider", "$httpProvider", "adalAuthenticationServiceProvider", function($routeProvider, $httpProvider, adalProvider) {
    $routeProvider.when("/login", {
        controller: "loginCtrl",
        templateUrl: "/app/templates/view-login.html",
        requireADLogin: false
    }).when ("/contacts", {
        controller: "contactsCtrl",
        templateUrl: "/app/templates/view-contacts.html",
        requireADLogin: true
    }).otherwise({
        redirectTo: "/login"
    });
    
    adalProvider.init({
        instance: "https://login.microsoftonline.com/",
        tenant: "rzdemos.com",
        clientId: "15f43fac-22db-4da6-9aa2-19037ea5138c",
        endpoints: {
            "MSGraph": "https://graph.microsoft.com"
        }
    }, $httpProvider);
}]);