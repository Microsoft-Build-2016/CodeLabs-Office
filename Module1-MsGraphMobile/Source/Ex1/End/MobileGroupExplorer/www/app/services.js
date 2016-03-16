(function () {
    "use strict";

    angular.module("myapp.services", []).factory("myappService", ["$rootScope", "$http", "$q", function ($rootScope, $http, $q) {
        var myappService = {};

        // Broadcasts waiting indicator toggles to app root
        myappService.wait = function (show) {
            $rootScope.$broadcast("wait", show);
        };

        // Broadcasts error messages to app root
        myappService.broadcastError = function (err) {
            $rootScope.$broadcast("error", err);
        };

        myappService.getGroups = function () {
            var deferred = $q.defer();

            getTokenForResource(graphResource).then(function (token) {
                // Use token to call into graph
                $http.defaults.headers.common["Authorization"] = "Bearer " + token.accessToken;
                $http.defaults.headers.post["accept"] = "application/json;odata=verbose";
                $http.get(graphResource + "/beta/me/joinedgroups").then(function (r) {
                    // Resolve the promise with group data
                    deferred.resolve(r.data.value);
                }, function (err) {
                    // Error calling API...reject the promise
                    deferred.reject("Groups failed to load");
                });
            }, function (err) {
                // Error getting token
                deferred.reject(err);
            });

            return deferred.promise;
        };

        myappService.loadPhotoAsync = function (obj, type) {
            var deferred = $q.defer();

            getTokenForResource(graphResource).then(function (token) {
                // Build the request url
                var url = graphResource + "/v1.0/" + type + "/" + obj.id + "/photo/$value";

                // Use token to call into graph
                $http.defaults.headers.common["Authorization"] = "Bearer " + token.accessToken;
                $http.defaults.headers.post["accept"] = "application/json;odata=verbose";
                $http.get(url, { responseType: "blob" }).then(function (image) {
                    // Convert blob into image that app can display
                    var imgUrl = window.URL || window.webkitURL;
                    var blobUrl = imgUrl.createObjectURL(image.data);
                    obj.img = blobUrl;
                    deferred.resolve();
                }, function (err) {
                    // Error calling API...reject the promise
                    deferred.reject("Image failed to load");
                });
            }, function (err) {
                // Error getting token
                deferred.reject(err);
            });

            return deferred.promise;
        };

        // Private resources
        var aadAuthContext = null;
        var aadAuthority = "https://login.microsoftonline.com/common";
        var aadAppClientId = "4b3e848c-111d-4b9c-94fd-ae03b78bc64a";
        var aadAppRedirect = "http://localhost:8000";
        var graphResource = "https://graph.microsoft.com";

        // Private function to get access token for a specific resource using ADAL
        var getTokenForResource = function (resource) {
            var deferred = $q.defer();

            getaadAuthContext().then(function (context) {
                // First try to get the token silently
                getTokenForResourceSilent(context, resource).then(function (token) {
                    // We were able to get the token silently...return it
                    deferred.resolve(token);
                }, function (err) {
                    // We failed to get the token silently...try getting it with user interaction
                    aadAuthContext.acquireTokenAsync(resource, aadAppClientId, aadAppRedirect).then(function (token) {
                        // Resolve the promise with the token
                        deferred.resolve(token);
                    }, function (err) {
                        // Reject the promise
                        deferred.reject("Error getting token");
                    });
                });
            });

            return deferred.promise;
        };

        // Private function to get access token for a specific resource silent using ADAL
        var getTokenForResourceSilent = function (context, resource) {
            var deferred = $q.defer();

            // read the tokenCache
            context.tokenCache.readItems().then(function (cacheItems) {
                // Try to get the roken silently
                var user_id;
                if (cacheItems.length > 1) {
                    user_id = cacheItems[0].userInfo.userId;
                }
                context.acquireTokenSilentAsync(resource, aadAppClientId, user_id).then(function (authResult) {
                    // Resolve the authResult from the silent token call
                    deferred.resolve(authResult);
                }, function (err) {
                    // Error getting token silent...reject the promise
                    deferred.reject("Error getting token silent");
                });
            }, function (err) {
                // Error getting cached data...reject the promise
                deferred.reject("Error reading token cache");
            });

            return deferred.promise;
        };

        // Private function to get authentication context using ADAL
        var getaadAuthContext = function () {
            var deferred = $q.defer();

            // Check if aadAuthContext is already initialized
            if (!aadAuthContext) {
                // aadAuthContext is null...initialize it
                aadAuthContext = Microsoft.ADAL.AuthenticationContext;
                aadAuthContext.createAsync(aadAuthority).then(function (context) {
                    // aadAuthContext is initialized...resolve in promise
                    aadAuthContext = context;
                    deferred.resolve(aadAuthContext);
                }, function (err) {
                    // aadAuthContext creation failed...reject the promise
                    aadAuthContext = null;
                    deferred.reject("Error creating auth context");
                });
            }
            else {
                // aadAuthContext is already initialized so resolve in promise
                deferred.resolve(aadAuthContext);
            }

            return deferred.promise;
        };

        return myappService;
    }]);
})();