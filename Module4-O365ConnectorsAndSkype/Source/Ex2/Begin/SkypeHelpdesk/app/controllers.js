angular.module("helpdesk.controllers", [])
.controller("loginCtrl", ["$scope", "$location", "adalAuthenticationService", function($scope, $location, adalSvc) {
    if (adalSvc.userInfo.isAuthenticated) {
        $location.path("/tickets");
    }
        
    $scope.login = function() {
        adalSvc.login();  
    };
}])
.controller("ticketsCtrl", ["$scope", "helpdeskSvc", function($scope, helpdeskSvc) {
    //get the helpdesk tickets
    helpdeskSvc.getTickets().then(function(tickets) {
        $scope.tickets = tickets; 
    });
    
    //get the user's profile picture
    $scope.pic = "/content/nopic.jpg";
    helpdeskSvc.getProfilePic().then(function(img) {
         $scope.pic = img;
    });
}]);