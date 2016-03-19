angular.module("helpdesk.controllers", [])
.controller("loginCtrl", ["$scope", "$location", "adalAuthenticationService", function($scope, $location, adalSvc) {
    if (adalSvc.userInfo.isAuthenticated) {
        $location.path("/tickets");
    }
        
    $scope.login = function() {
        adalSvc.login();  
    };
}])
.controller("ticketsCtrl", ["$scope", "helpdeskSvc", "skypeSvc", function($scope, helpdeskSvc, skypeSvc) {
    //get the helpdesk tickets
    helpdeskSvc.getTickets().then(function(tickets) {
        $scope.tickets = tickets; 
        
        //ensure the user is signed into skype
        skypeSvc.ensureSignIn().then(function() {
            angular.forEach($scope.tickets, function(ticket, index) {
                //look up the status of the user and listen for changes
                skypeSvc.subscribeToStatus(ticket.created_by.email).then(function(status) {
                    ticket.created_by.status = status;
                    if (!$scope.$$phase)
                        $scope.$apply(); 
                });
            });
        });
    });
    
    //get the user's profile picture
    $scope.pic = "/content/nopic.jpg";
    helpdeskSvc.getProfilePic().then(function(img) {
         $scope.pic = img;
    });
    
    //listen for status changes
    $scope.$on("statusChanged", function(evt, data) {
         var id = data.user.id().replace("sip:", "");
       
        //find all instances of this user
        angular.forEach($scope.tickets, function(ticket, index) {
            if (ticket.created_by.email === id) {
                ticket.created_by.status = data.status;
                if (!$scope.$$phase)
                    $scope.$apply(); 
            }
        });
    });
    
    //helper function to check if a status can perform chat
    var canChat = function(status) {
        var chattableStatus = { 
            Online: true, Busy: true, Idle: true, IdleOnline: true, Away: true, BeRightBack: true,
            DoNotDisturb: false, Offline: false, Unknown: false, Hidden: false };
        return chattableStatus[status];
    };
    
    //starts a chat
    $scope.startChat = function(ticket) {
        if (canChat(ticket.created_by.status)) {
            skypeSvc.startConversation("sip:" + ticket.created_by.email);
            $scope.showChatWindow = true;
        }
    };
    
    //closes the chat window
    $scope.closeChatWindow = function() {
        $scope.showChatWindow = false;
    };
}]);