<a name="HOLTop" />
# Web Development and the Microsoft Graph #

---

<a name="Overview" />
## Overview ##

The **Microsoft Graph** provides a single developer end-point accessing  Microsoft cloud services. The **Microsoft Graph** can be used to develop application against both commercial and consumer services, including **OneDrive**, **OneDrive for Business**, **Outlook.com**, **Exchange Online**, **OneNote**, **SharePoint Online**, **Azure AD**, and much more. By delivering multiple services through one end-point, the **Microsoft Graph** exposes the relationships that exist between different Microsoft services and makes it easy for developers to traverse these relationships.

For example, you can use the **Microsoft Graph** to look up users (**Azure AD**), their mail/calendar (**Exchange Online**), the manager they report to (**Azure AD**), the groups they collaborate in (**Exchange** & **SharePoint Online**), and even the files they collaborate on with their manager (**OneDrive for Business**).

Azure Active Directory (**Azure AD**) is the identity provider that secures the Microsoft Graph. With it, developers can register applications that connect to the Microsoft Graph and its underlying services. The Azure AD team offers libraries and plugins to make authentication against Azure AD easy from almost any platform. The Azure Active Directory Authentication Libraries (**ADAL**) as offered for all major mobile platforms, including Windows, iOS, Android, Xamarin, and Apache Cordova.

This module will show you how to build a powerful and secure web application that leverages **Azure AD** for identity and calls into the **Microsoft Graph**. It will cover both client-side and server side scenarios.

<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:

- Register web applications in Azure Active Directory
- Learn about both organizational and consumer scenarios for leveraging the Microsoft Graph
- Build client-side applications that connect to the Microsoft Graph using an implicit flow
- Develop server-side applications with the Microsoft Graph and converged authentication

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater

[1]: https://www.visualstudio.com/products/visual-studio-community-vs

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [AngularJS and the Microsoft Graph](#Exercise1)
2. [ASP.NET Core and the Microsoft Graph](#Exercise2)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise1"></a>
### Exercise 1: AngularJS and the Microsoft Graph ###

The popularity of client-side applications and JavaScript frameworks have been on a tremendous surge. However, these applications also introduce some significant challenges such as client-side authentication flows and cross-domain API calls. Luckily, **Azure AD** offers an **Implicit OAuth Flow** for client-side applications, and the **Microsoft Graph** supports cross-origin resource sharing (aka - **CORS**) for secure cross-domain API calls. Together these can be used to build powerful Office 365 applications with most client-side framework.

In this exercise, you will build a **single-page application** (**SPA**) using **AngularJS** to connect to the **Microsoft Graph**. You will walk through application registry, leveraging **ADAL.js**, and calling into the **Microsoft Graph**.

<a name="Ex1Task1"></a>
#### Task 1 - Registering the Application with Azure AD ####

In this task, you'll go through the steps of registering an app in **Azure AD** using the **Office 365 App Registration Tool**. You can also register apps in the **Azure Management Portal** ([https://manage.windowsazure.com](https://manage.windowsazure.com "https://manage.windowsazure.com")), but the **Office 365 App Registration Tool** allows you to register applications without the need for an Azure Subscription. **Azure AD** is the identity provider for Office 365, so any service/application that wants to use Office 365 data must be registered in Azure AD.

1. Open a browser and navigate to the **Office 365 App Registration Tool** at **[http://dev.office.com/app-registration](http://dev.office.com/app-registration "http://dev.office.com/app-registration")**.

2. The **Office 365 App Registration Tool** welcome screen will give you the option to use an existing Office 365 tenant or create a new Office 365 tenant. Click the **Sign in with your Office 365 account** and use your Office 365 account or the one that was provided to you.

	![App registration sign-in](Images/Mod2_AppRegSignin.png?raw=true "App registration sign-in")

	_App registration sign-in_

3. Once you are signed in, the app registration page will allow you to specify the details of your application. Use the details outlined below and then click the **Register App** button:

	- **App name**: **Angular Contacts**
	- **App type**: **Web**
	- **Sign on URL**: **http://localhost:8000**
	- **Redirect URI**: **http://localhost:8000**
	- **App permissions**: **Contacts.Read**

4. Once the app registration is complete, the app registration page should display registration confirmation that includes a **Client ID** (below the **Register App** button). Copy this **Client ID** (or leave the browser up) for use in the next exercise. The registration tool will also list a **Client Secret**, but you can ignore that as we will be using an **Implicit OAuth Flow** that doesn't use the Secret.

	![Registration Confirmation](Images/Mod2_AppRegConfirm.png?raw=true "Registration Confirmation")

	_Registration Confirmation_

<a name="Ex1Task2"></a>
#### Task 2 - Leveraging Adal-Angular ####

This exercise uses a starter project with basic project scaffolding pre-configured. Nothing special, just a single-page application that uses **AngularJS** and **Angular Routing**. In this exercise, you will import the **ADAL-Angular** module and leverage it assist in authentication based on the Angular routes that are defined in the application.  

1. Open command prompt and change directory to the module's **Source\Ex1\Begin** folder. This is the starter folder for this exercise.

2. Open the web project in **Visual Studio Code** by typing **code .**

		code .

3. The application contains three areas of interest for this exercise:

	- **index.html**: the single html that will host all the applications content.
	- **app**: the folder containing all of the application logic and partial views/templates.
	- **lib**: the folder containing all the frameworks/dependent scripts (ex: Bootstrap, Angular, etc). All of these were imported using **bower**.

	![Project structure](Images/Mod2_ProjStruct.png?raw=true "Project structure")

	_Project structure_

4. Open the **app.js** file located in the **app** folder. Notice that the **o365Service.getContacts** function returns hard-coded contact data. This will be converted to real contacts from Office 365 by the end of the exercise.

5. Return to the command prompt (still set to the project start folder) and start a static web server by typing **superstatic --port 8000**.

	````CMD
	superstatic --port 8000
	````

	![superstatic](Images/Mod2_ss.png?raw=true "superstatic")

	_superstatic_

6. Open a browser and navigate to **http://localhost:8000**. Get a feel for the flow of the application. It has two views...a **login** view and a **list** view that display contacts. Currently, the application is hard-coded, so the sign-in button on the login form just switches views to a hard-coded list of contacts.

	![Hard-coded contacts](Images/Mod2_ContactsHardCode.png?raw=true "Hard-coded contacts")

	_Hard-coded contacts_

7. Let's convert the application to leverage the application registration we completed in Task 1. First, stop the static web server by typing **Ctrl-C** in the command window.

8. Next, we need to import the **Azure Active Directory Authentication Library** (**ADAL**) using **Bower**. On the command prompt type bower install adal-angular --save.

	````CMD
	bower install adal-angular --save
	````

9. If you return to Visual Studio Code and expand the **lib** folder, you should now see the **adal-angular** library that **Bower** imported.

10. In Visual Studio Code, open the **index.html** file in the root of the project.

11. Add script references to **adal.min.js** and **adal-angular.min.js** after the **angular** and **angular-route** reference. Both of these scripts are located in the **lib\adal-angular**:

	````XML
	<!-- ADAL references -->
	<script type="text/javascript" src="lib/adal-angular/dist/adal.min.js"></script>
	<script type="text/javascript" src="lib/adal-angular/dist/adal-angular.min.js"></script>
	````

12. Next, open the **app.js** file located in the **app** folder. This is the file that contains all of the application logic.

13. Locate the **myContacts** Angular module that is defined towards the bottom of the file. Modify it to have a dependency on **AdalAngular** and the **config** function to leverage the **adalAuthenticationServiceProvider**. Both of these are examples of AngularJS use of **dependency injection**. That is, we are injecting the **AdalAngular** module and **adalAuthenticationServiceProvider** object into **myContacts**.

	````JavaScript
	angular.module("myContacts", ["myContacts.services",
		"myContacts.controllers", "ngRoute", "AdalAngular"])
		.config(["$routeProvider", "$httpProvider", "adalAuthenticationServiceProvider",
			function($routeProvider, $httpProvider, adalProvider) {
	````

14. Inside the **config** function, notice how **$routeProvider** is used to configure the application routes. There are two defined routes (**login** and **contacts**) and an otherwise block to handle any other undefined route. Modify each of the routes with a boolean **requireADLogin** attribute. This attribute is made possible by the **AdalAngular** module and will force the route to be signed in when set to true. The **login** route should be **requireADLogin** to **false** and **contacts** should set **requireADLogin** to **true**;

	````JavaScript
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
		}]);
	````

15. Next, use the **adalProvider.init()** function to define the app details from the app registration you performed in Task 1. You can add this by using the **o365-adalInit** code snippet directly after the **$routeProvider** defines the routes. You will need to fill in values for **tenant** and **clientId**.

	````JavaScript
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
			tenant: "mytenant.onmicrosoft.com",
			clientId: "15f43fac-22db-4da6-9aa2-19037ea5138c",
			endpoints: {
				"MSGraph": "https://graph.microsoft.com"
			}
		}, $httpProvider);
	}]);
	````

16. Return to the command prompt (still set to the project start folder) and start a static web server by typing **superstatic --port 8000**.

	````CMD
	superstatic --port 8000
	````

17. Open a browser and navigate to **http://localhost:8000**. This time when you click on the **Sign-in with Office 365** button, the application should force you to sign-in with Office 365. This is because the contacts route requires Azure AD Login (via **requireADLogin** attribute).

	![Sign-in](Images/Mod2_SignIn.png?raw=true "Sign-in")

	_Sign-in_

18. You have successfully implemented Azure AD authentication using **ADAL.js**. In the next exercise, you will replace the hard-coded contact data with real contacts in the **Microsoft Graph**.


<a name="Ex1Task3"></a>
#### Task 3 - Calling the Microsoft Graph ####

In the final Task of Exercise 1, you will convert the hard-coded contacts to real data from **Office 365** using the **Microsoft Graph**.

1. In Visual Studio Code, open the **app.js** file in the **app** folder.

2. Update the **loginCtrl** of the **myContacts.controllers** module to leverage the **adalAuthenticationService** dependency and use it to change the application flow based on if the user is authenticated (**adalSvc.userInfo.isAuthenticated**).

	````JavaScript
	angular.module("myContacts.controllers", [])
	.controller("loginCtrl", ["$scope", "$location",
		"adalAuthenticationService",
		function($scope, $location, adalSvc) {

		if (adalSvc.userInfo.isAuthenticated) {
			$location.path("/contacts");
		}

		$scope.login = function() {
			adalSvc.login();  
		};
	}])
	````

3. Finally, update the **o365Service.getContacts** function in the **myContacts.services** module to use the **$http** object to **GET** the signed-in user's contacts using the **Microsoft Graph** end-point **https://graph.microsoft.com/v1.0/me/contacts**.

	````JavaScript
	o365Service.getContacts = function() {
		var deferred = $q.defer();

		$http.get("https://graph.microsoft.com/v1.0/me/contacts").then(function(result) {
			deferred.resolve(result.data);
		}, function(err) {
			deferred.reject(err.statusText);
		});

		return deferred.promise;
	}
	````

4. Every call into the Microsoft Graph requires an **access token** passed as an **Authorization** header of the request. You would normally set these headers on the **$http** object in the above code. However, the **AdalAngular** module automatically does this for you. In fact, **ADAL** handles all of the token management (caching, headers, etc) for you!

5. Open a browser and navigate to **http://localhost:8000**. After clicking **Sign-in with Office 365** the application should query the Microsoft Graph and display actual contacts from Office 365. If you want to test further, go to [https://outlook.office365.com](https://outlook.office365.com "https://outlook.office365.com") to add a few more contacts.

6. This completes Exercise 1. Hopefully the exercise helped illustrate how easy it is to convert and AngularJS application to leverage the **Microsoft Graph**.

<a name="Exercise2"></a>
### Exercise 2: ASP.NET Core and the Microsoft Graph ###

Exercise 1 concentrated on an enterprise development scenario where organizational accounts queried contacts in Exchange Online using the Microsoft Graph. What if you wanted your app to query contacts in a consumer service like Outlook.com? In the past, this would require completely separate app registrations, authentication flows, and API end-points. However, Azure AD and the Microsoft Graph support a new application model that converges the developer experience for commercial and consumer services.

The v2.0 "**converged**" app model has a few notable differences:

- One application definition can support both web and native platforms and consumer (MSA) and commercial (O365) users.
- Permissions are passed in dynamically at runtime in a scope parameter
- Refresh tokens aren't automatic...you have to request one via offline_access permission scope
- Resources are inferred by the permission scopes passed in

In this exercise you will work with the new v2.0 app model by building an ASP.NET Core application that connects to the Microsoft Graph for both consumer and commercial users.

> **Important Note:** ASP.NET Core, Azure AD v2.0 end-points, and the Microsoft Graph end-points that use v2.0 are very new and do not support all scenarios and familiar libraries. This exercise uses some experimental helper libraries that will change. I repeat...**some of the code in this exercise WILL change**. It is advised that you concentrate on the high-level concepts of v2.0 "converged" auth and not the library details.

<a name="Ex2Task1"></a>
#### Task 1 - Registering the App in new App Portal ####

The new v2.0 "converged" application model uses a centralized registration portal that both **Azure AD** and **MSA** accounts can use to register applications (read: no requirement for an Azure Subscription). In this Task, you will use the new **Application Registration Portal** to register a v2.0 application. Later, you will use this application to query contacts in Exchange Online or Outlook.com (depending on the account used to sign-in).

1. Open a browser and navigate to [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com "https://apps.dev.microsoft.com").

	![Application Registration Portal](Images/Mod2_AppPortal.png?raw=true "Application Registration Portal")

	_Application Registration Portal_

2. Click the **Sign in with Microsoft** button.

3. On the **Sign In** screen, you can use either an organization account (ex: Office 365) or Microsoft account (ex: Outlook.com, Live.com, etc). For this exercise, use your **Office 365 account** or the one that was assigned to you.

4. Once you sign into the Application Registration Portal, click the **Add an app** to launch the **New Application Registration** dialog.

	![New app registration dialog](Images/Mod2_AppRegV2.png?raw=true "New app registration dialog")

	_New app registration dialog_

5. Provide a **Name** for the new app and click the **Create application** button.

6. On the new app confirmation screen, click the **Generate New Password** button to generate an application password. Make sure you copy this value before closing the **New password generated** dialog...you display it again.

	![Generate secret](Images/Mod2_Secret.png?raw=true "Generate secret")

	_Generate secret_

7. Once you have copied and closed the New password generated dialog,  copy the **Application Id** GUID.

8. scroll down to the **Platforms** section and click the **Add Platform** button and then select **Web** for the new platform.

	![Add platform](Images/Mod2_platforms.png?raw=true "Add platform")

	_Add platform_

9. When the Add Platform dialog closes, enter the **Redirect URI** of **https://localhost:44300/signin-oidc** and then click the Save button at the bottom of the screen.

	![app redirect](Images/Mod2_Redirect.png?raw=true "app redirect")

	_app redirect_


10. The v2.0 application is registered and ready to be used in the next tasks. If you have registered applications with Azure AD in the past, you will notice that no permissions needed to be pre-configured. The v2.0 model passes permission scopes in when requesting tokens. We will explore that more in the subsequent tasks.

<a name="Ex2Task2"></a>
#### Task 2 - Configuring Authentication ####

You will leverage your new app registration in an ASP.NET Core web application. ASP.NET Core is still in preview, so it and the libraries it uses are subject to (and probably will) change.

1. Launch **Visual Studio 2015** and launch the new Project dialog (**File**\**New**\**Project**).

2. In the new project dialog select the **ASP.NET Web Application** template under the **Visual C#** > **Web** templates.

	![New project](Images/Mod2_NewProj1.png?raw=true "New project")

	_New project_

3. On the **New ASP.NET Project** dialog select **Web Application** in the **ASP.NET 5 Templates** section and then click the **Change Authentication** button.

	![Web project type](Images/Mod2_NewProj2.png?raw=true "Web project type")

	_Web project type_

4. On the **Change Authentication** dialog, change the option to **No Authentication**. You will manually add authentication using an OWIN startup class.

	![No Auth](Images/Mod2_NoAuth.png?raw=true "No Auth")

	_No Auth_

5. Once you have changed the authentication, click the **OK** button to provision the new project using the ASP.NET 5 template.

6. To get started in the new project, import the following NuGet packages using the Package Manager Console.

	````CMD
	Install-Package Microsoft.AspNet.Authentication.OpenIdConnect -Pre
	Install-Package Microsoft.AspNet.Authentication.Cookies -Pre
	Install-Package Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory -Pre
	Install-Package Microsoft.AspNet.Session -Pre
	````

	> **Important Note:** Just to emphasize again, these libraries are prerelease packages and will likely change. The ADAL library is also marked "experimental" and is a preview library for working with v2.0 apps. It will also change as it comes closer to release.

7. Next, you need to configure the app to use SSL and the port you specified in the app registration. Right-click the web project and select **Properties** to bring up the project properties screen.

8. Select the Debug link in the left navigation. And scroll down to the **Web Server Settings** section.

9. Check the **Enable SSL** checkbox and change the **App URL** to **https://localhost:44300** before saving and closing the project properties screen.

	![Project Properties](Images/Mod2_ProjProp.png?raw=true "Project Properties")

	_Project Properties_

10. Next, open the appsettings.json file and add a AzureAD section with config values for AppId, AppPassword, Tenant, Authority, and GraphResourceId.

	````JavaScript
	{
		"Logging": {
			"IncludeScopes": false,
			"LogLevel": {
				"Default": "Verbose",
				"System": "Information",
				"Microsoft": "Information"
			}
		},
		"AzureAD": {
			"AppId": "YOUR_APPLICATION_ID",
			"AppPassword": "YOUR_APP_PASSWORD",
			"Tenant": "YOUR_TENANT.onmicrosoft.com",
			"AadInstance": "https://login.microsoftonline.com/",
			"GraphResourceId": "https://graph.microsoft.com"
		}
	}
	````

11. Next, create a **Utils** folder in the root of the web project and then a **SettingsHelper.cs** class in the **Utils** folder.

12. Use the **o365-settingshelper** code snippet to populate the **SettingsHelper** class and resolve missing references (ex: **using Microsoft.Extensions.Configuration**). This class reads configuration settings you set in Step 10 and makes them available as strongly-typed properties.

13. Next, create a **SessionTokenCache.cs** class in the **Utils** folder and set it to inherit from **TokenCache** (in the **Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory** namespace).

	````C#
	using Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory;
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Threading.Tasks;

	namespace MyContactsV2App.Utils
	{
		public class SessionTokenCache : TokenCache
		{
		}
	}
	````


15. Next, use the **o365-sessiontokencache** code snippet to populate the remainder of the **SessionTokenCache** class. You may need to resolve a reference to **Microsoft.AspNet.Http**. This class will provide the token caching mechanism for **ADAL**.

16. Finally, open the **Startup.cs** file in the root of the web project. After **Line 51** (app.UseStaticFiles()), insert the **o365-startup** code snippet. You will also have to resolve the following references.

	````C#
	using Microsoft.AspNet.Authentication.Cookies;
	using Microsoft.IdentityModel.Protocols.OpenIdConnect;
	using Microsoft.AspNet.Authentication.OpenIdConnect;
	using Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory;
	````

17. The **o365-startup** code snippet configured **session state**, **cookie authentication**, and **OpenIdConnect** authentication with the v2.0 application model.

18. Look at the **UseOpenIdConnectAuthentication** method and notice its use of **Scope** to specify the permissions for the app (**Contacts.ReadWrite** and **offline_access**). Also notice the use of **OnAuthorizationCodeReceived** to handle authorization codes coming back from a user's sign in and app consent.

	````C#
	app.UseOpenIdConnectAuthentication(options =>
	{
		options.AutomaticChallenge = true;
		options.ClientId = SettingsHelper.AppId;
		options.Authority = SettingsHelper.Authority;
		options.SignInScheme = CookieAuthenticationDefaults.AuthenticationScheme;
		options.ResponseType = OpenIdConnectResponseTypes.CodeIdToken;
		options.Scope.Add("https://graph.microsoft.com/contacts.readwrite");
		options.Scope.Add("offline_access");
		options.Events = new OpenIdConnectEvents
		{
			OnAuthorizationCodeReceived = async (context) =>
			{
				string userObjectId = context.AuthenticationTicket.Principal.FindFirst(SettingsHelper.ObjectIdentifierKey).Value;
				ClientCredential clientCred = new ClientCredential(SettingsHelper.AppId, SettingsHelper.AppPassword);
				AuthenticationContext authContext = new AuthenticationContext(options.Authority, false, new SessionTokenCache(userObjectId, context.HttpContext));
				AuthenticationResult authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
					context.Code, new Uri(context.RedirectUri), clientCred, new string[] { "https://graph.microsoft.com/contacts.readwrite" });
			}
		};
	});
	````

18. While still in the **Startup.cs** file, locate the **ConfigureServices** method and update it as seen below.

	````C#
	// This method gets called by the runtime. Use this method to add services to the container.
	public void ConfigureServices(IServiceCollection services)
	{
		// Add framework services.
		services.AddMvc();
		services.AddSession();
		services.AddCaching();
	}
	````

19. Authentication has been configured, but you need to update a controller to enforce it. Open the **HomeController.cs** file in the **Controllers** folder and add the [**Authorize**] directive to the Index action.

	````C#
	[Authorize]
	public IActionResult Index()
	{
		return View();
	}
	````

20. It is time to test your work. Press **F5** or start the debugger. When the application loads, it should bring up a sign-in screen that you can provide either a Office 365 or Microsoft (MSA) account on. After sign-in, you will be presented with a consent screen to authorize the app the permission scopes passed in. This consent screen will look slightly different if using an MSA account, but it achieves the same thing.

	![Consent](Images/Mod2_Consent.png?raw=true "Consent")

	_Consent_

21. After consenting the application, the browser should return to the application. Did it work? Yes, as long as you are on the default view. In the next Task, you will mark REST calls into the **Microsoft Graph**.

	![Authenticated](Images/Mod2_Authenticated.png?raw=true "Authenticated")

	_Authenticated_


<a name="Ex2Task3"></a>
#### Task 3 - Calling the Microsoft Graph ####

I the final exercise of this module, you will use a v2.0 access token and call into the Microsoft Graph for both Office 365 and Microsoft (MSA) accounts.

1. Open the **Index.cshtml** located in the **Views**\**Home** folder.

2. Replace the entire view markup with the **o365-contactsindex** code snippet. This configures a table to display the contacts. It also defines the model of the view as a **Newtonsoft.Json.Linq.JArray**. This has been simplified for the workshop. A more solid pattern would be to create a Contact class and parse the JSON from the **Microsoft Graph**.

	> **Note:** Microsoft is working to release a number of SDKs for the Microsoft Graph. At the time of this writing, a preview SDK was available for ASP.NET MVC 4.x, but not yet supported in ASP.NET Core. Stay posted for updated on SDKs that can simplify data access with the Microsoft Graph.

3. Open the **HomeController.cs** file in the **Controllers** folder and add update the entire Index action with the **o365-contactsindex** code snippet.

	````C#
	[Authorize]
	public async Task<IActionResult> Index()
	{
		JArray jsonArray = null;

		// Get access token for calling into Microsoft Graph
		string userObjectId = ((ClaimsIdentity)User.Identity).Claims.FirstOrDefault(i => i.Type == SettingsHelper.ObjectIdentifierKey).Value;
		ClientCredential clientCredential = new ClientCredential(SettingsHelper.AppId, SettingsHelper.AppPassword);
		AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, false, new SessionTokenCache(userObjectId, HttpContext));
		var token = await authContext.AcquireTokenSilentAsync(new string[] { "https://graph.microsoft.com/contacts.readwrite" }, clientCredential, UserIdentifier.AnyUser);

		// Use the token to call Microsoft Graph
		HttpClient client = new HttpClient();
		client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token.Token);
		client.DefaultRequestHeaders.Add("Accept", "application/json");
		using (var response = await client.GetAsync(SettingsHelper.GraphResourceId + "/v1.0/me/contacts"))
		{
			if (response.IsSuccessStatusCode)
			{
				var json = await response.Content.ReadAsStringAsync();
				JObject jObj = JObject.Parse(json);
				jsonArray = jObj.Value<JArray>("value");
			}
		}

		return View(jsonArray);
	}
	````

4. You will likely need to resolve a number of reference after inserting this snippet.

	````C#
	using Newtonsoft.Json.Linq;
	using Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory;
	using System.Net.Http;
	using System.Security.Claims;
	````

5. Here are a few important things to analyze in this snippet.

	- The SessionTokenCache calls is use to retrieve tokens
	- The access token is being set in the Authorization header as a Bearer token
	- The Microsoft Graph uses a simple me/contacts path to look up the users path

6. It is time to test your work. Press **F5** or start the debugger. After signing into the application, the Home view should display contacts for the user that signed in. This works with NO code change between Office 365 and Microsoft (MSA) accounts!

	![Completed solution](Images/Mod2_complete.png?raw=true "Completed solution")

	_Completed solution_

<a name="Summary" />
## Summary ##

By completing this module, you should have:

- Registered a web applications in Azure Active Directory
- Explored both organizational and consumer scenarios for leveraging the Microsoft Graph
- Built a client-side application that connect to the Microsoft Graph using an implicit flow
- Developed a server-side application with the Microsoft Graph and converged authentication
