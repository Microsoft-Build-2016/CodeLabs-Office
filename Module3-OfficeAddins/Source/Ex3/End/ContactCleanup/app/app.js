/* global $ */
var app = (function() {  // jshint ignore:line
    "use strict";

    var self = {};
    
    self.clientId = "4acded18-b0d2-4f90-9bec-c011517964d0";
    self.tenant = "rzdemos.com";
    self.redirectUri = "https://localhost:8443/app/auth/auth.html";

    // Common initialization function (to be called from each page)
    self.initialize = function() {
        $("body").append(
            "<div id='notification-message'>" +
            "<div class='padding'>" +
            "<div id='notification-message-close'></div>" +
            "<div id='notification-message-header'></div>" +
            "<div id='notification-message-body'></div>" +
            "</div>" +
            "</div>");

        $("#notification-message-close").click(function() {
            $("#notification-message").hide();
        });

        // After initialization, expose a common notification function
        self.showNotification = function(header, text) {
            $("#notification-message-header").text(header);
            $("#notification-message-body").text(text);
            $("#notification-message").slideDown("fast");
        };
    };

    return self;
})();
