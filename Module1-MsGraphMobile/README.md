<a name="HOLTop" />
# Mobile Development and the Microsoft Graph #

---

<a name="Overview" />
## Overview ##

Microsoft offers powerful tools for creating mobile applications that connect to secure services in the cloud. One of the most interesting services is the **Microsoft Graph**, which provides a single end-point to securely access most of the commercial and consumer services Microsoft offers, including **OneDrive**, **OneDrive for Business**, **Outlook.com**, **Exchange Online**, **OneNote**, **SharePoint Online**, **Azure AD**, and much more.

Azure Active Directory (**Azure AD**) is the identity provider that secures the Microsoft Graph. With it, developers can register applications that connect to the Microsoft Graph and its underlying services. The Azure AD team offers libraries and plugins to make authentication against Azure AD easy from almost any platform. The Azure Active Directory Authentication Libraries (**ADAL**) as offered for all major mobile platforms, including Windows, iOS, Android, Xamarin, and Apache Cordova.

This module will show you how to build a powerful and secure mobile application that leverages Azure AD for identity and calls into the Microsoft Graph.

<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:

- Register a mobile (aka - "native") application in Azure Active Directory
- Leverage the Azure Active Directory Authentication Library (ADAL) plug-in for Apache Cordova to easily authenticate mobile users against Azure AD
- Create and test a mobile app that connects to the Microsoft Graph and works on most modern mobile platforms using Apache Cordova
- Work with Office 365 group data and photos

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater
- [Visual Studio Tools for Apache Cordova][2]
- [Visual Studio Emulator for Android][3]

[1]: https://www.visualstudio.com/products/visual-studio-community-vs
[2]: https://www.visualstudio.com/en-us/features/cordova-vs.aspx
[3]: https://www.visualstudio.com/en-us/features/msft-android-emulator-vs.aspx

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [Getting Familiar with the Solution](#Exercise1)
2. [Registering the App with Azure AD](#Exercise2)
3. [Implementing the ADAL plug-in for Authentication](#Exercise3)
4. [Using Microsoft Graph to load Office 365 Groups](#Exercise4)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise1"></a>
### Exercise 1: Getting Familiar with Solution ###

In this module you will build a native mobile application using the **Apache Cordova**, which is framework for building cross-platform mobile applications using client-side web development. Apache Cordova apps are installed like other native apps and have a rich ecosystem of plug-ins to perform advanced device-specific operations. The Visual Studio Tools for Apache Cordova allow Apache Cordova apps to be built and tested in Visual Studio.

This exercise uses a starter project with basic project scaffolding pre-configured. The solution uses the **Ionic Framework** and **AngularJS** to facilitate the app design and functionality. Ionic is an open source front-end SDK that makes it incredibly easy to build beautiful and interactive mobile apps with Apache Cordova. You do not need any prior experience with Apache Cordova, Ionic, or AngularJS to complete this module.

<a name="Ex1Task1"></a>
#### Task 1 - Exploring the Solution ####
In this tasks you will explore the starter solution in Visual Studio and get familiar with the project structure of a project leveraging **Apache Cordova**, **Ionic**, and **AngularJS**.

1. Open Windows Explorer and browse to the module's **Source\Ex1\Begin** folder.

2. Double-click the solution file (**MobileGroupExplorer.sln**) to open the solution in **Visual Studio Community 2015**.

3. Bring up **Solution Explorer** to familiarize yourself with the project scaffolding.

	![Starter project in Solution Explorer](Images/Mod1_SlnExplorer.png?raw=true "Starter project in Solution Explorer")

	_Starter project in Solution Explorer_

4. Open and explore the **config.xml** file, which contains the configuration for the app, including app details (name, icons, splash screens) and plug-ins. Click on the different links in the left navigation to get an idea of the different configuration settings available.

5. The **www** folder contains all the web artifacts that make up the user interface of an Apache Cordova application. Expand this folder and open the **Index.html** file. This application is considered a **Single-Page Application** (**SPA**). A SPA displays a single HTML page (**Index.html**) in which partial views are dynamically loaded into. In this application, the dynamic partial views are loaded into the **ion-nav-view** directive on **Line 16**.

	````HTML
	<body>
		<div class="bar bar-header bar-assertive" ng-controller="appCtrl" ng-show="activeError" style="z-index: 1000">
			<h1 class="title">Error: {{error}}</h1>
			<button class="button icon ion-close-round" ng-click="dismissError()"></button>
		</div>
		<ion-nav-view></ion-nav-view>
		<div class="spinner" ng-controller="appCtrl" ng-show="wait">
			<ion-spinner icon="android" class="spinnerSpinner"></ion-spinner>
			<div class="spinnerBackground"></div>
		</div>
	````

6. The partial views are located under **www\app\templates**.

7. Open and explore the **index.js** file located under **www\scripts**. This is the default script file that the Apache Cordova project template includes. It subscribes to device events such as **deviceready**, **pause**, and **resume**. On **Line 18**, we are attaching an **AngularJS** module named to **myapp** to the **Index.html** page (also called "**bootstrapping**" the module).

	````JavaScript
	(function () {
		"use strict";

		document.addEventListener( 'deviceready', onDeviceReady.bind( this ), false );

		function onDeviceReady() {
			// Handle the Cordova pause and resume events
			document.addEventListener( 'pause', onPause.bind( this ), false );
			document.addEventListener( 'resume', onResume.bind( this ), false );

			// TODO: Cordova has been loaded. Perform any initialization that requires Cordova here.

			// Bootstrap the root module to the page
			angular.bootstrap(document, ["myapp"]);
		};
	````

8. Most of the app functionality is contained in the **services.js**, **controllers.js**, and **app.js** files in the **www\app** folder. Let's start by opening and exploring the **services.js** file. This module  (**myapp.services**) defines a **factory** which provides persistent data and functionality across the app. Notice right now this factory has group data hard-coded. By the end of the module, this will be dynamic data queried from the **Microsoft Graph**.

	````JavaScript
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
	````

9. Open and explore the **controllers.js** file located under **www\app**. **Ionic** and **AngularJS** make use of a **Model**/**View**/**Controller** (**MVC**) architecture, and this module (**myapp.controllers**) contains all the controllers for the application. Notice controllers for **appCtrl** and **homeCtrl**.

10. Open and explore the **app.js** file located under **www\app**. This file defines the primary module (**myapp**) for the application. **AngularJS** make heavy use of **dependency injection**. The brackets following the module name defines the other modules **myapp** is dependent on, including **Ionic**, **myapp.controllers**, and **myapp.services**.

	````JavaScript
	(function () {
		"use strict";

		angular.module("myapp", ["ionic",
			"myapp.controllers", "myapp.services"])
	````

11. **app.js** also defines all the routes that make the app more dynamic. Notice that each route defined the **url** (path), **controller**, and **templateUrl** (view) that is used on the route. The route table also defines an **otherwise** criteria, which is used when an unrecognized path is encountered in the application.

	````JavaScript
	.config(function ($stateProvider, $urlRouterProvider) {
		$stateProvider
		.state("app", {
			url: "/app",
			abstract: true,
			templateUrl: "app/templates/view-menu.html",
			controller: "appCtrl"
		})
		.state("app.home", {
			url: "/home",
			templateUrl: "app/templates/view-home.html",
			controller: "homeCtrl"
		});
		$urlRouterProvider.otherwise("/app/home");
	});
	````

<a name="Ex1Task2"></a>
#### Task 2 - Debugging Apache Cordova ####
In this task, you will explore the debugging options for an Apache Cordova project in Visual Studio Community 2015 and debug the solution using the Visual Studio Emulator for Android.

1. Look at the debugging toolbar in Visual Studio Community 2015 and locate the **Solution Platform** and **Debug Target** options.

	![Solution Platform and Debug Tagets](Images/Mod1_Targets.png?raw=true "Solution Platform and Debug Tagets")

	_Solution Platform and Debug Tagets_

2. The **Solution Platform** dropdown allows you to select a platform target. It should have options for **Android**, **iOS**, **Windows Phone (Universal)**, **Windows Phone 8**, **Windows-AnyCPU**, **Windows-ARM**, **Windows-x64**, and **Windows-x86**. Each **Solution Platform** has its own unique **Debug Target** options. Look at some of the different **Debug Target** options by selecting different **Solution Platforms**. Note that most of the options allow for both emulator and physical device debugging options.

3. Select **Android** as the **Solution Platfrom** and **VS Emulator 5" KitKat (4.4) XXHDPI Phone** as **Debug Target**.

4. Press **F5** or start the debugger to launch the **Visual Studio Emulator for Android**.

5. Once the emulator launches and Android boots, the application should automatically launch with the Visual Studio debugger attached. Right now the application doesn't do much as it has hard-coded group data. This completes **Exercise 1**, but you might keep the emulator up for subsequent exercises.

	![Start project with hard-coded groups](Images/Mod1_Debug1.png?raw=true "Start project with hard-coded groups")

	_Start project with hard-coded groups_

>**Note**: It might take a few minutes for the emulator to completely boot up Android. If the User Account Control dialog box is shown for Hyper-V related escalation, confirm the action to proceed.

<a name="Exercise2"></a>
### Exercise 2: Registering the App with Azure AD ###

Now that you are familiar with the project structure, it's time to convert the application to authenticate against Azure Active Directory and call into the Microsoft Graph. In this exercise, you'll register an application in Azure AD and capture app details that will be used in subsequent exercises.

<a name="Ex2Task1"></a>
#### Task 1 - Registering the App ####

In this task, you'll go through the steps of registering an app in Azure AD using the **Office 365 App Registration Tool**. You can also register apps in the **Azure Management Portal**, but the **Office 365 App Registration Tool** allows you to register applications without having access to the **Azure Management Portal**.

1. Open a browser and navigate to the **Office 365 App Registration Tool** at **[http://dev.office.com/app-registration](http://dev.office.com/app-registration "http://dev.office.com/app-registration")**.

2. The **Office 365 App Registration Tool** welcome screen will give you the option to use an existing Office 365 tenant or create a new Office 365 tenant. Click the **Sign in with your Office 365 account** and use your Office 365 account or the one that was provided to you.

	![Office 365 App Registration Tool sign-in](Images/Mod1_AppRegSignin.png?raw=true "Office 365 App Registration Tool sign-in")

	_Office 365 App Registration Tool sign-in_

3. Once you are signed in, the app registration page will allow you to specify the details of your application. Use the details outlined below and then click the **Register App** button:

	- **App name**: **MobileGraphExplorer**
	- **App type**: **Native**
	- **Redirect URI**: **http://localhost:8000**
	- **App permissions**: **Groups.Read.All**

4. Once the app registration is complete, the app registration page should display registration confirmation that includes a **Client ID** (below the **Register App** button). Copy this **Client ID** (or leave the browser up) for use in the next exercise.

	![App registration confirmation with Client ID](Images/Mod1_AppRegConfirm.png?raw=true "App registration confirmation with Client ID")

	_App registration confirmation with Client ID_

<a name="Exercise3"></a>
### Exercise 3: Implementing the ADAL plug-in for Authentication ###

In this exercise, you'll take the application registration from Exercise 2 and leverage it with the Apache Cordova mobile app in Visual Studio. This involves adding the Azure Active Directory Authentication Library (ADAL) plug-in and writing some script against ADAL in our project.

<a name="Ex3Task1"></a>
#### Task 1 - Adding the ADAL plug-in ####

In this task, you'll add the Azure Active Directory Authentication Library (ADAL) plug-in for Apache Cordova to the project.

1. Locate and open the **config.xml** file in the root of the project to open the Apache Cordova project configuration screen.

2. Click on the **Plugins** link in the side navigation of the configuration screen. The **Plugins** view allows you to add powerful plug-ins to your Apache Cordova project. There are a number of popular "**Core**" plug-ins listed, but you can also import local plug-ins or plug-ins from Git using the **Custom** option. Locate the **Active Directory Authentication Library (ADAL)** plugin and click the **Add** button.

	![ADAL Cordova plug-in](Images/Mod1_ADAL.png?raw=true "ADAL Cordova plug-in")

	_ADAL Cordova plug-in_

3. Once the **ADAL** plug-in finishes downloading, open the **services.js** file located in the **www\app** folder.

4. Right before the return statement on **Line 31**, insert the app settings code snippet by typing **o365-appsettings** and pressing **tab**. This will insert some private variable into the module that contain details about your app. You should update the value **aadAppClientId** to the **Client ID** you received when registering your application in **Exercise 2** (also update **aadAppRedirect** if you did not go with the default value of **http://localhost:8000**).

	````JavaScript
	var aadAuthContext = null;
	var aadAuthority = "https://login.microsoftonline.com/common";
	var aadAppClientId = "e712e8f5-c813-4c0d-b3fa-9e44d184e168";
	var aadAppRedirect = "http://localhost:8000";
	var graphResource = "https://graph.microsoft.com";
	````

5. Next, we need to add some script that uses the **ADAL** plug-in to assist with authentication and getting tokens for calling the Microsoft Graph. Directly under the settings you just added to the **service.js** file, add the ADAL Cordova code snippet by typing **o365-adalcordova** and pressing **tab**.

6. The **o365-adalcordova** code snippet adds a number of functions that makes authentication and token acquisition a breeze. Here is a high-level summary of the functions (also note that all three functions return promises):

	- **getTokenForResource**: gets an access token for a specific resource. First tries to get the token silently from token cache (using getTokenForResourceSilent) but prompts the user if silent acquisition fails. This is the function that is called before making calls to the Microsoft Graph.
	- **getTokenForResourceSilent**: attempts to get an access token for a  specific resource silently by using the token cache for the application. Called by getTokenForResource
	- **getaadAuthContext**: returns the authentication context (instantiating it if necessary).

7. In the same **services.js** file, locate the **getGroups** function (**Line 24**) and modify it to call **getTokenForResource** before resolving the promise. You can pass the **graphResource** variable as the resource parameter in **getTokenForResource**:

	````JavaScript
	myappService.getGroups = function () {
		var deferred = $q.defer();

		getTokenForResource(graphResource).then(function (token) {
			//Ignore token for now and resolve the promise
			deferred.resolve(groups);
		});

		return deferred.promise;
	};
	````

8. Press **F5** or start the debugger to deploy the updated app to the **Visual Studio Emulator for Android**.

9. When the app loads, it should display an Office 365 sign-in screen. You didn't have to build a sign-in screen...ADAL provided you with one. Sign-in using your Office 365 account or the one that was provided to you.

	![Sign-in via ADAL plug-in](Images/Mod1_Signin.png?raw=true "Sign-in via ADAL plug-in")

	_Sign-in via ADAL plug-in_

10. After you sign-in, the you will be prompted to consent to the permissions the mobile app is requesting (that you configured when you registered the app). Click the **Accept** button to consent to the application permissions.

	![Consent via ADAL plug-in](Images/Mod1_Consent.png?raw=true "Consent via ADAL plug-in")

	_Consent via ADAL plug-in_

11. After consenting the application, the ADAL will close all the Azure AD screens and return to the home view of the application. The groups displayed are still hard-coded, but that will change in **Exercise 4**.

	> **Note:** Once you consent to the application, you won't be prompted again for consent (unless you delete the consent). You also won't be prompted to sign-in next time you launch the app, as the ADAL plug-in automatically handles token caching. You should only be prompted sign-in after your token invalidates (ex: Office 365 password changes), you go 14 days without launching the app, or after 90 days of continual use.


<a name="Exercise4"></a>
### Exercise 4: Using the Microsoft Graph to load Office 365 Groups ###

In this exercise, you'll take the token you acquired in Exercise 3 and use it to call into the Microsoft Graph to query data stored in Office 365. Specifically, you will query the new Office 365 groups and group photos.

<a name="Ex4Task1"></a>
#### Task 1 - Querying Office 365 Groups with the Microsoft Graph ####

In this task, you'll replace the hard-coded groups with actual group data from Office 365 by calling the Microsoft Graph.

1. Open the **services.js** file located in the **www\app** folder.

2. Add the **$http** dependency on factory definition.

	````JavaScript
	(function () {
		"use strict";

		angular.module("myapp.services", [])
		.factory("myappService", ["$rootScope", "$http", "$q",
			function ($rootScope, $http, $q) {
				var myappService = {};
				...
	````

3. Locate the **getGroups** function and update it to use **http** to **GET** the current user's joined groups with the Microsoft Graph [graph.microsoft.com/beta/me/joinedgroups](https://graph.microsoft.com/beta/me/joinedgroups). This operation should be completed directly after the call to **getTokenForResource**.

	````JavaScript
	myappService.getGroups = function () {
		var deferred = $q.defer();

		getTokenForResource(graphResource).then(function (token) {
			// Use token to call into graph
			$http.get(graphResource + "/beta/me/joinedgroups").then(function (r) {
				// Resolve the promise with group data
				deferred.resolve(r.data.value);
			});
		});

		return deferred.promise;
	};
	````

4. Calling into the Microsoft Graph requires that the access token is attached to the **Authorization** header of all request (as **Bearer** token). Update the code to set this header and the **accept** header to **application/json;odata=verbose**.

	> **Note:** The code below also implements error checking on the getTokenForResource and $http calls. If an error occurs, the getGroups function rejects the promise.

	````JavaScript
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
	````

5. Next, open the **controllers.js** file located at **www\app** and update the code after **Line 49** to default an **img** property for each group to **"images/group.png"**. In **Task 2**, we will query the Microsoft to display the actual group photo.

	````JavaScript
	myappService.getGroups().then(function (data) {
		// Refresh binding
		$scope.groups = data;
		$scope.$broadcast("scroll.refreshComplete");
		myappService.wait(false);

		// Lazy load the group photos
		angular.forEach($scope.groups, function (obj, key) {
			obj.img = "images/group.png";
		});
	}, function (err) {
		// Stop the refresh and broadcast the error
		myappService.broadcastError(err);
		$scope.$broadcast("scroll.refreshComplete");
		myappService.wait(false);
	});
	````

6. Press **F5** or start the debugger to deploy the updated app to the **Visual Studio Emulator for Android**.

7. When the app loads, it should display a waiting indicator before loading the Office 365 groups you are a member of. All groups should display the same generic group photo, but in **Task 2** you will update the app the lazy load the actual group photos.

	![App with groups using Microsoft Graph](Images/Mod1_Debug2.png?raw=true "App with groups using Microsoft Graph")

	_App with groups using Microsoft Graph_

<a name="Ex4Task2"></a>
#### Task 2 - Lazy Load Group Photos with the Microsoft Graph ####

In this task, you'll update the solution to lazy load group photos by calling the Microsoft Graph.

1. Open the **services.js** file located at **www\app**.

2. Directly below the **getGroups** function, define a **loadPhotoAsync** function on the **myappService** that takes **obj** and **type** parameters. The **obj** parameter will be object we are loading a photo for and the **type** is the type of object in the Microsoft Graph (ex: users, groups, etc). Also configure the new function to return a **promise**.

	````JavaScript
	myappService.loadPhotoAsync = function (obj, type) {
		var deferred = $q.defer();

		return deferred.promise;
	};
	````

3. The Microsoft Graph uses the following end-point pattern to get photos: **https://graph.microsoft.com/{version}/{type}/{objectid}/photo/$value**. Hopefully you can see that the **obj** and **type** parameters are all we need to get photos for groups or users. Update the **loadPhotoAsync** function to **GET** the photo for the group passed in. Remember that you need a token (via **loadPhotoAsync**) and it must be included in the **Authorization** header of the request.

	````JavaScript
	myappService.loadPhotoAsync = function (obj, type) {
		var deferred = $q.defer();

		getTokenForResource(graphResource).then(function (token) {
			// Build the request url
			var url = graphResource + "/v1.0/" + type + "/" + obj.id + "/photo/$value";

			// Use token to call into graph
			$http.defaults.headers.common["Authorization"] = "Bearer " + token.accessToken;
			$http.defaults.headers.post["accept"] = "application/json;odata=verbose";
			$http.get(url).then(function (image) {
				//TODO: work with image
			});
		});

		return deferred.promise;
	};
	````

4. Update the **loadPhotoAsync** method to handle errors on the **getTokenForResource** and **$http** calls.

	````JavaScript
	myappService.loadPhotoAsync = function (obj, type) {
		var deferred = $q.defer();

		getTokenForResource(graphResource).then(function (token) {
			// Build the request url
			var url = graphResource + "/v1.0/" + type + "/" + obj.id + "/photo/$value";

			// Use token to call into graph
			$http.defaults.headers.common["Authorization"] = "Bearer " + token.accessToken;
			$http.defaults.headers.post["accept"] = "application/json;odata=verbose";
			$http.get(url).then(function (image) {

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
	````

5. Working with photos in the Microsoft Graph are a little from typical GETs. For starters, update the $http.get operation to indicate a **responseType** of **blob**.

	````JavaScript
	$http.get(url, { responseType: "blob" }).then(function (image) {
		...
	````

6. Setting the **responseType** to **blob** will allow the **$http** object to handle to photo blob coming back from the Microsoft Graph. This blob can't immediately be used to set the src attribute of an image in HTML. Instead, we need to add some code to create an object URL.

	````JavaScript
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
	````

7. Finally, we need to update the **homeCtrl** in the **controller.js** file to call our new **loadPhotoAsync** function. You can call this in the loop we created in **Step 5** of the previous Task.

	````JavaScript
	myappService.getGroups().then(function (data) {
		// Refresh binding
		$scope.groups = data;
		$scope.$broadcast("scroll.refreshComplete");
		myappService.wait(false);

		// Lazy load the group photos
		angular.forEach($scope.groups, function (obj, key) {
			obj.img = "images/group.png";
			myappService.loadPhotoAsync(obj, "groups").then(function () {
			});
		});
	}, function (err) {
		// Stop the refresh and broadcast the error
		myappService.broadcastError(err);
		$scope.$broadcast("scroll.refreshComplete");
		myappService.wait(false);
	});
	````

8. Press **F5** or start the debugger to deploy the updated app to the **Visual Studio Emulator for Android**.

9. When the app loads, it should display a waiting indicator before loading the Office 365 groups you are a member of. The groups will initially load with the generic group.png image, but will later load the actual group photos from Office 365.

	![Groups with group photos loaded through Microsoft Graph](Images/Mod1_Debug3.png?raw=true "Groups with group photos loaded through Microsoft Graph")

	_Groups with group photos loaded through Microsoft Graph_

<a name="Summary" />
## Summary ##

By completing this module, you should have:

- Registered a mobile (aka - "native") application in Azure Active Directory
- Leveraged the Azure Active Directory Authentication Library (ADAL) plug-in for Apache Cordova to easily authenticate mobile users against Azure AD
- Created and tested a mobile app that connects to the Microsoft Graph and works on most modern mobile platforms using Apache Cordova
- Worked with Office 365 group data and photos
