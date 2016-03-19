angular.module("helpdesk", ["helpdesk.services", "helpdesk.controllers", "ngRoute", "AdalAngular"])
.config(["$routeProvider", "$httpProvider", "adalAuthenticationServiceProvider", function($routeProvider, $httpProvider, adalProvider) {
    $routeProvider.when("/login", {
        controller: "loginCtrl",
        templateUrl: "/app/templates/view-login.html",
        requireADLogin: false
    }).when ("/tickets", {
        controller: "ticketsCtrl",
        templateUrl: "/app/templates/view-tickets.html",
        requireADLogin: true
    }).otherwise({
        redirectTo: "/login"
    });
    
    adalProvider.init({
        instance: "https://login.microsoftonline.com/",
        tenant: "TENANT.onmicrosoft.com", //TODO: CHANGE THIS TO YOUR OFFICE 365 TENANT DOMAIN
        clientId: "6fd45769-7a1e-4dc5-a876-90fa781b3d3e",
        endpoints: {
            "https://webdir.online.lync.com": "https://webdir.online.lync.com",
            "https://graph.microsoft.com": "https://graph.microsoft.com"
        }
    }, $httpProvider);
}]);