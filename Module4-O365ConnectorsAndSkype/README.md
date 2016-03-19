<<<<<<< HEAD
<a name="HOLTop" />
# Integrating Conversations with Skype and Office 365 Connectors #

---

<a name="Overview" />
## Overview ##

Office 365 now enables developers to build a powerful new generation of **conversation**-driven applications. **Office 365 Connectors** and **Skype SDKs** are two developer platforms to lighting up conversations in your applications. Office 365 Connectors allow 3rd party applications to send data into Office 365 and Outlook in near real-time. Users can subscribe to relevant data/applications and have informative and actionable "cards" display in Outlook. Skype SDKs allow developers to being the power capabilities of Skype into their applications built for the web and mobile platforms.

In this module, you will explore development with Office 365 Connectors and Skype SDKs, including how to integrate them into existing web applications.

<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:

- Manually register an Office 365 Connector and use a webhook to send data into Office 365.
- Connect an existing web application to Office 365 via Office 365 Connectors, including the addition of "Connect to Office 365" buttons, handling callbacks from Office 365, and sending data into Office 365 via webhooks.
- Skype-enable and existing web application using the Skype Web SDK.

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater
- [Microsoft Office 2016][2]
- [Any Web Request Composer][2]

[1]: https://www.visualstudio.com/products/visual-studio-community-vs
[2]: https://portal.office.com
[3]: https://www.hurl.it

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.

<a name="Setup" />
### Setup ###
In order to run the exercises in this module, you'll need to set up your environment first.

1. Open Windows Explorer and browse to the module's **Source** folder.
2. Right-click **Setup.cmd** and select **Run as administrator** to launch the setup process that will configure your environment and install the Visual Studio code snippets for this module.
3. If the User Account Control dialog box is shown, confirm the action to proceed.

> **Note:** Make sure you've checked all the dependencies for this module before running the setup.

<a name="CodeSnippets" />
### Using the Code Snippets ###

Throughout the module document, you'll be instructed to insert code blocks. For your convenience, most of this code is provided as Visual Studio Code Snippets, which you can access from within Visual Studio 2015 to avoid having to add it manually.

>**Note**: Each exercise is accompanied by a starting solution located in the **Begin** folder of the exercise that allows you to follow each exercise independently of the others. Please be aware that the code snippets that are added during an exercise are missing from these starting solutions and may not work until you've completed the exercise. Inside the source code for an exercise, you'll also find an **End** folder containing a Visual Studio solution with the code that results from completing the steps in the corresponding exercise. You can use these solutions as guidance if you need additional help as you work through this module.

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [Building Office 365 Connectors](#Exercise1)
2. [Developing with the Skype Web SDK](#Exercise2)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise1"></a>
### Exercise 1: Building Office 365 Connectors ###

Office 365 Connectors are a great way to get useful information and content into your Office 365 Group. Any user can connect their group to services like Trello, Bing News, Twitter, etc., and get notified of the group's activity in that service. From tracking a team's progress in Trello, to following important hashtags in Twitter, Office 365 Connectors make it easier for an Office 365 group to stay in sync and get more done. Office 365 Connectors also provides a compelling extensibility solution for developers which we will explore in this Exercise.

In this Exercise, you will explore the developer options for working with Office 365 Connectors, including adding connectors, leveraging webhooks, and integrating Office 365 Connectors into existing web application via the "Connect to Office 365" button.

<a name="Ex1Task1"></a>
#### Task 1 - Getting Started with Office 365 Connectors ####

In this Task, you will explore Office 365 Groups and some of the existing Office 365 Connectors that are available.

1. Open a browser and navigate to [https://portal.office.com](https://portal.office.com "https://portal.office.com") and sign-in with the Office 365 credentials that were provided to you.

2. **Office 365 Connectors** for Groups are currently under developer preview. In order to access them, navigate to **OWA** using the URL [https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true](https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true "https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true")

3. Once you are signed into OWA, locate the Office 365 Groups you are a member of in the lower left navigation.

	![Groups in OWA](http://i.imgur.com/YJwx2ai.png)

4. Create your own unique group by clicking the **+** (plus) sign to the right of the **Groups** title in the left navigation.

	![Add Group](http://i.imgur.com/UvRDOAK.png)

5. Provide a **name** and **description** for the new group and click **Create** (optionally add members once the groups has been created).

6. Select **More > Connectors** from the Group's top navigation (if More isn't an option, you might try navigating back to the group using the link in **Step 2**).

	![Connectors in group navigation](http://i.imgur.com/5D6xEQi.png)

7. Explore some of the Office 365 Connectors that are available out of the box.

	![OOTB Connectors](http://i.imgur.com/ARrQVET.png)

8. Locate the **Twitter** connector and click **Add**.

9. The Twitter connector requires you to sign-in with a Twitter account. To do this, click on the Log in button.

	![Twitter Connector Log in](http://i.imgur.com/wjaXiHh.png)

10. After authorizing the Office 365 Connector for Twitter, you can select specific **users**/**hashtags** to follow and how **frequently** they show up in the Office 365 Group. Try to follow yourself or a hashtag.

	![Configure Twitter Connector](http://i.imgur.com/Zb6Pesd.png)

11. Post a tweet that matches the criteria in Step 10 and see it show up in the Office 365 Group.

	![Twitter post via Office 365 Connector](http://i.imgur.com/UId0Seg.png)

<a name="Ex1Task2"></a>
#### Task 2 - Leveraging Webhooks with Office 365 Connectors ####

Hopefully Task 1 helped to illustrate the power of Office 365 Connectors, but did little to showcase the unique developer opportunity. In this Task, you will explore Office 365 Connector webhooks and how developer can leverage then to send data into Office 365.

1. Navigate to the Office 365 Group you created in the previous Task and select **More > Connectors** from the Group's top navigation.

2. Locate the **Incoming Webhook** Connector and click the **Add** button.

	![Incoming Webhook](http://i.imgur.com/oC4uDam.png)

3. Specify a **name** for the incoming webhook (ex: Build 2016) and click the **Create** button.

4. The confirmation screen will display a **URL** that is the webhook end-point we will use later in this Task.

	![Webhook Confirmation](http://i.imgur.com/N7PJsxC.png)

5. Open a new browser tab and navigate to [https://www.hurl.it](https://www.hurl.it "https://www.hurl.it"), which is an in-browser web request composer similar to what Fiddler offers.

6. When the page loads, add the following details:
	- **Operation**: **POST**
	- **Destination Address**: **webhook URL** from **Step 4**
	- **Headers**: **Content-Type: application/json** 
	- **Body**: **{ "text": "Hello from Build 2016" }**

	![Manual Webhook](http://i.imgur.com/vV8FKeD.png)

7. Accept the **Captcha** and click the **Launch Request** button. You should get a confirmation screen that looks similar to the following.

	![Webhook Manual Confirmation](http://i.imgur.com/LjEi7m6.png)

8. If you return to the Office 365 Group, you should be able to locate the message you sent into it via the webhook.

	![Message sent into Group via webhook](http://i.imgur.com/VtFJkTQ.png)

9. Although you sent a very simple message into the webhook, Office 365 Connectors support a much more complex message format. You can get more details on the message format by visiting [https://dev.outlook.com/Connectors/GetStarted](https://dev.outlook.com/Connectors/GetStarted "https://dev.outlook.com/Connectors/GetStarted").

<a name="Ex1Task3"></a>
#### Task 3 - Integreting "Connect to Office 365" into Existing Applications ####

Task 2 had you manually register a webhook for an Office 365 Connector. In this Task, you will modify an existing web application to register webhooks with Office 365 by leveraging a "Connect to Office 365" button. You will capture the webhook details in a custom callback and send messages to Office 365 when new records are created in the application.

This Task uses a starter project to serve as the existing application. The application is a Craigslist-style selling site named BillsList. You are tasked with enhancing BillsList to allow users to subscribe to listing categories and send messages to Office 365 Groups when new listings match the subscription criteria. 

1. Open Windows Explorer and browse to the module's **Source > Ex1 > Begin** folder.

2. Double-click the solution file (**BillsListASPNET.sln**) to open the solution in **Visual Studio Community 2015**.

3. The starter solution actually has two project...**BillsListASPNET** (the web application) and **BillsListASPNET.Data** (database project). Right-click the **BillsListASPNET.Data** project and select **Publish**.

4. On the **Publish Database** dialog, click on the **Edit** button to configure the connection information.

	![Publish DB](http://i.imgur.com/jyCqnDd.png)

5. On the **Connection Properties** dialog enter **(localdb)\MSSQLLocalDB** for the **Server name** and click **OK**.

	![Connection Properties](http://i.imgur.com/0urRfhM.png)

6. When you return to the **Publish Database** dialog, click the **Publish** button to publish the database to **LocalDb**.

7. When the database has finished publishing, press **F5** or start the debugger to test the starter project.

8. When the application loads, click on **Listings** in the top navigation. This will prompt you to sign-in. Use the **Office 365 account** that was provided to you (it also supports Consumer/MSA accounts like outlook.com, live.com, hotmail.com, etc).

9. The **Listings** view is one of the views we want to modify to support subscriptions to Office 365 Connectors. Notice that listings also have a category link, which is the second view we will add the "**Connect to Office 365**" button.

	![Listings](http://i.imgur.com/mx5SdfB.png)
 
10. Close the browser to stop debugging and open the **Index.cshtml** file located in the web project at **Views > Items**.

11. Copy and paste the following markup at the end of the **H2** element (right after the text "**for sale**").

        <a style="float: right;" href="https://outlook.office.com/connectors/ConnectToO365?state=@Request.Url.AbsoluteUri&app_name=BillsList&app_logo_url=http://billslist.azurewebsites.net/images/logo_128.png&callback_url=https://localhost:44300/callback">
            <img src="https://connecttoo365.blob.core.windows.net/images/ConnectToO365Button.png" alt="Connect to Office 365" style="height: 32px;"></img>
        </a>    

12. This snippet will add a "Connect to Office 365" button to the view. It passes the following parameters to Office 365:

	- **state**: optional state information...in our case we are passing the current view information so we can return to it after the connection has been established with Office 365
	- **app_name**: the name of the application (ex: BillsList)
	- **app_logo_url**: the application logo that will be displayed in Office 365 when messages are sent in via webhook
	- **callback**: the location that Office 365 will return webhook information to after the user has confirmed the connection (ex: https://localhost:44300/callback)

13. Next, open the **Category.cshtml** file located in the web project at **Views > Items**.

14. Copy and paste the following markup at the end of the **H2** element (right after "**@ViewData["category"]**").

        <a style="float: right;" href="https://outlook.office.com/connectors/ConnectToO365?state=@Request.Url.AbsoluteUri&app_name=BillsList%20(@ViewData["category"])&app_logo_url=http://billslist.azurewebsites.net/images/logo_128.png&callback_url=https://localhost:44300/callback">
            <img src="https://connecttoo365.blob.core.windows.net/images/ConnectToO365Button.png" alt="Connect to Office 365" style="height: 32px;"></img>
        </a>

15. This markup snippet is slightly different. It sends a dynamic **app_name** to Office 365 that includes the category the user is subscribing to. This allows the user to subscribe to specific categories instead of ALL listings.

16. You might recall we are passing in a callback location of **https://localhost:44300/callback** to Office 365. However, the **Callback** controller does not yet exist...let's create it. Right click the **Controllers** folder in the web project and select **Add > Controller**.

17. Select **MVC Controller - Empty** for the controller type and name it **CallbackController**.

	![New Controller](http://i.imgur.com/oEfHEsy.png)

18. Inside the **CallbackController** class, add the **o365-callbackctrl** code snippet by typing **o365-callbackctrl** and pressing **tab**.

        // GET: Callback
        public ActionResult Index()
        {
            var error = Request["error"];
            var state = Request["state"];
            if (!String.IsNullOrEmpty(error))
            {
                return RedirectToAction("Error", "Home", null);
            }
            else
            {
                var group = Request["group_name"];
                var webhook = Request["webhook_url"];
                Subscription sub = new Subscription();
                sub.GroupName = group;
                sub.WebHookUri = webhook;

                //set optional category
                if (state.IndexOf("?c=") != -1)
                    sub.Category = state.Substring(state.IndexOf("?c=") + 3);

                //save the subscription
                using (BillsListEntities entities = new BillsListEntities())
                {
                    entities.Subscriptions.Add(sub);
                    entities.SaveChanges();
                    return Redirect(state);
                }
            }
        }

19. This controller looks for information returned from Office 365 and saves it as a subscription. The specific information passed from Office 365 as parameters include:

	- **error**: error details if the connection with Office 365 failed (ex: user rejected the connection)
	- **state**: the state value that was passed in via the "Connect to Office 365" button. In our case, it could include the category the user subscribed to
	- **group_name**: the name of the group the user selected to connect to
	- **webhook_url**: the webhook end-point our application will use to send messages into Office 365

20. Almost done, just need to update the **Create** activity to send messages to the appropriate webhooks. Open the **ItemsController.cs** file located in the **Controllers** folder of the web project.

21. Towards the bottom of the class, add the **o365-callwebhook** code snippet by typing **o365-callwebhook** and pressing **tab** (resolve using dependencies if necessary).

        private async Task callWebhook(string webhook, Item item)
        {
            var imgString = "https://billslist.azurewebsites.net/images/logo_40.png";
            if (Request.Files.Count > 0)
            {
                //resize the image
                Request.Files[0].InputStream.Position = 0;
                Image img = Image.FromStream(Request.Files[0].InputStream);
                var newImg = (Image)(new Bitmap(img, new Size(40, 40)));

                //convert the stream
                using (var stream = new System.IO.MemoryStream())
                {
                    newImg.Save(stream, ImageFormat.Jpeg);
                    stream.Position = 0;
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    imgString = "data:image/jpg;base64, " + Convert.ToBase64String(bytes);
                }
            }

            //prepare the json payload
            var json = @"
                {
                    'summary': 'A new listing was posted to BillsList',
                    'sections': [
                        {
                            'activityTitle': 'New BillsList listing',
                            'activitySubtitle': '" + item.Title + @"',
                            'activityImage': '" + imgString + @"',
                            'facts': [
                                {
                                    'name': 'Category',
                                    'value': '" + item.Category + @"'
                                },
                                {
                                    'name': 'Price',
                                    'value': '$" + item.Price + @"'
                                },
                                {
                                    'name': 'Listed by',
                                    'value': '" + item.Owner + @"'
                                }
                            ]
                        }
                    ],
                    'potentialAction': [
                        {
                            '@context': 'http://schema.org',
                            '@type': 'ViewAction',
                            'name': 'View in BillsList',
                            'target': [
                                'https://localhost:44300/items/detail/" + item.Id + @"'
                            ]
                        }
                    ]}";

            //prepare the http POST
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
            using (var response = await client.PostAsync(webhook, content))
            {
                //TODO: check response.IsSuccessStatusCode
            }
        }

22. This snippet takes the new listing details and sends it to Office 365 via **POST** to the webhook end-point.

23. Finally, locate the **Create** activity within the class. **Create** is overloaded, so select the one that is marked with **HttpPost** and has the **Item** parameter. Inside the using statement add the code below between **SaveChanges()** of the new listing and the **RedirectToAction()** statement. This identifies matching subscriptions and calls the appropriate webhooks.

        //save the item to the database
        using (BillsListEntities entities = new BillsListEntities())
        {
            entities.Items.Add(item);
            var id = entities.SaveChanges();

            //loop through subscriptions and call webhooks for each
            var subs = entities.Subscriptions.Where(i => i.Category == null || i.Category == item.Category);
            foreach (var sub in subs)
            {
                await callWebhook(sub.WebHookUri, item);
            }

            return RedirectToAction("Detail", new { id = item.Id });
        }

24. It's time to test your work. Press **F5** or start the debugger to launch the application. When you click on **Listings** (and sign-in) the **Listings** view will have a "**Connect to Office 365**" button in the upper right. This button will also display on the **Category** view.

	![Connect to Office 365 button](http://i.imgur.com/1lsNgEP.png)

25. View a specific **Category** and then click on the "**Connect to Office 365**" button. You should be redirected to a screen to select a Office 365 Group to connect to.

	![Select group to connect to](http://i.imgur.com/FyzHqKJ.png)

26. Select an Office 365 Group and click **Allow** to complete establish the connection with Office 365 and return to BillsList.

27. To test the connection, click on **My Listings** and **Create listing** with the **category** you subscribed to. A Connector **Card** for the listing should almost immediately show up in the Office 365 Group.

![Connector card](http://i.imgur.com/xfTGxOZ.png)

<a name="Exercise2"></a>
### Exercise 2: Developing with the Skype Web SDK ###

Skype is one of the most popular communication platforms in the world. Many organization look to **Skype for Business** to deliver their real-time communication needs. Skype for Business offers powerful SDKs to integrate real-time conversations into both web and mobile application.

In this exercise, you will convert and existing web application to integrate Skype **presence** and **instant messaging** with the **Skype Web SDK**. You will also see how to integrate additional conversation modalities like **voice** and **video**.

<a name="Ex2Task1"></a>
#### Task 1 - Setup ####

This exercise uses a starter solution that follows a help desk scenario.  It displays help desk tickets assigned to the signed in user. Your job in the exercise is to integrate Skype presence and instant messaging into the help desk application. The starter solution is built with AngularJS and is already configured to authenticate against Azure AD and display the user's profile picture using the Microsoft Graph. Although it helps, you do not need prior experience with AngularJS...basic JavaScript and HTML skills will suffice.

In the first task, you will familiarize yourself with the starter solution and get it running locally.

1. Open a command prompt and browse to the module's **Source > Ex1 > Begin** folder.

2. Open the solution by typing code .

		code .

3. The starter solution is considered an AngularJS single page application (SPA) because it uses single index.html page to host all the content. It leverages a Model/View/Controller model to load dynamic content in the single page. The list below lists some of the significant components of the starter solution.

	- **index.html**: the single html that will host all the applications content.
	- **app**: the folder containing all of the application logic and partial views/templates.
		- **templates**: contains all the HTML partial views for the application.
		- **app.js**: defines the primary Angular module for the application and the application routes.
		- **factory.js**: defines an Angular factory that provides properties and services across the application.
		- **controllers.js**: defines all the controllers for the application.
	- **lib**: the folder containing all the frameworks/dependent scripts (ex: Bootstrap, Angular, etc). All of these were imported using **bower**.

4. You should be given **two Office 365 accounts** for this exercise. You should identify one as the **Skype User** and one as the **Web User**. The Skype User will sign into **Skype for Business** and the Web User will use the web application built in this exercise. Open the **factory.js** file in the **app** folder. Go to **lines 32-33** and update the **tenantDomain** and **skypeTestUser** with the tenant domain and Skype User respectively. The example below updates these settings with the contoso tenant and john user.

        31		//Hack...will use hard-coded tickets for demo purposes
        32		var tenantDomain = "contoso.onmicrosoft.com"; //CHANGE THIS TO YOUR OFFICE 365 TENANT
        33		var skypeTestUser = "john@contoso.onmicrosoft.com"; //CHANGE THIS TO THE USER THAT WILL TEST FROM SKYPE


5. Open the **app.js** file in the **app** folder. Notice how each route includes an attribute to determine if authentication is required or not. This is enabled through **ADAL-Angular**, which is the **Azure Active Directory Authentication Library** (**ADAL**) that manages authentication for the application.

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

6. Locate the section of **app.js** where the **ADAL** settings are configured and update the **tenant** property with the tenant domain you are using in Office 365.

        adalProvider.init({
            instance: "https://login.microsoftonline.com/",
            tenant: "TENANT.onmicrosoft.com", //TODO: CHANGE THIS TO YOUR OFFICE 365 TENANT DOMAIN
            clientId: "6fd45769-7a1e-4dc5-a876-90fa781b3d3e",
            endpoints: {
                "https://webdir.online.lync.com": "https://webdir.online.lync.com",
                "https://graph.microsoft.com": "https://graph.microsoft.com"
            }
        }, $httpProvider);

7. Open the **controllers.js** file in the **app** folder and locate the **loginCtrl**. Notice it's use of **adalSvc** to check if the user is authenticated. 

        .controller("loginCtrl", ["$scope", "$location", "adalAuthenticationService", function($scope, $location, adalSvc) {
            if (adalSvc.userInfo.isAuthenticated) {
                $location.path("/tickets");
            }
                
            $scope.login = function() {
                adalSvc.login();  
            };
        }])

8. Return to the command prompt and type superstatic --port 8000. This will start a simple web server to host your client-side web application.

		superstatic --port 8000

9. Open a browser and navigate to **http://localhost:8000**. The site should direct you to the **login** view and prompt you to sign-in with **Office 365**.

	![Sign-in with Office 365](http://i.imgur.com/kLy3oRb.png)

10. Sign into the web application using the credentials of the Web User. Once you sign-in, the application should display a list of help desk tickets. In the next Task, you will modify this view to display Skype presence for each user.

	![Help desk tickets](http://i.imgur.com/9Qha0e6.png)

<a name="Ex2Task2"></a>	
#### Task 2 - Sign-in and Presence with Skype for Business ####

In this task, you will introduce the Skype Web SDK into the solution and use it to subscribe and display presence for users in the help desk application. Applications build against the Skype Web SDK need to be registered in Azure AD. We will demonstrate this process, but you will use a predefined application ID in this lab.

1. Create a **skype.js** file in the **app** folder of the solution and populate it with base scaffolding using the **o365-skypefactory** code snippet. This creates a **skype.services** Angular module and populates it with some of the core settings to integrate the **Skype Web SDK**. The **skypeSvc** factory will provide persistent objects and services across all controllers of the application. It uses a **singleton** pattern for defining properties a functions.

        angular.module("skype.services", [])
        .factory("skypeSvc", ["$rootScope", "$http", "$q", function($rootScope, $http, $q) {
            var skypeSvc = {};

            //private properties
            var apiManager = null;
            var client = null;

            //config settings for the app
            skypeSvc.config = {
                apiKey: "a42fcebd-5b43-4b89-a065-74450fb91255", // SDK DF
                apiKeyCC: "9c967f6b-a846-4df2-b43d-5167e47d81e1", // SDK+CC DF
                initParams: {
                    auth: null,
                    client_id: "6fd45769-7a1e-4dc5-a876-90fa781b3d3e", //Client ID of app in Azure AD
                    cors: true,
                    origins: ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
                    redirect_uri: "/auth.html",
                    version: "sdk-samples/1.0.0" // this helps to identify telemetry generated by the samples
                }
            };

            //Add additional Skype logic here...signin, status, conversations, etc

            return skypeSvc;
        }]);

2. Take note of the redirect_uri property above that is set the /auth.html. This is a page the Skype Web SDK will open in a hidden iFrame to help with token acquisition.

3. Next, open the **index.html** file in the root of the project and add a script reference to the **Skype Web SDK** below the bootstrap reference (you can also use the **o365-skyperef** code snippet for this).

        <!-- JQuery and Bootstrap references -->
        <script type="text/javascript" src="lib/jquery/dist/jquery.min.js"></script>
        <script type="text/javascript" src="lib/bootstrap/dist/js/bootstrap.min.js"></script>
        
        <!-- Skype reference -->
        <script src="https://swx.cdn.skype.com/shared/v/latest/SkypeBootstrap.min.js"></script>

4. You also need to add a reference to the **skype.js** Angular factory you created in step 1. Add that in the app scripts section between the **factory.js** and the **controllers.js**.

        <!-- App scripts -->
        <script type="text/javascript" src="app/factory.js"></script>
        <script type="text/javascript" src="app/skype.js"></script>
        <script type="text/javascript" src="app/controllers.js"></script>
        <script type="text/javascript" src="app/app.js"></script>

5. Next, you need to inject a dependency of the **skype.services** module you created in step 1 in the main Angular module of the application. Open the app.js file and modify the helpdesk module definition as follows.

		angular.module("helpdesk", ["skype.services", "helpdesk.services", "helpdesk.controllers", "ngRoute", "AdalAngular"])

6. Return to the **skype.js** file you created in Step 1. In the "Add additional Skype logic here" section, add an **ensureClient** function on the **skypeSvc** object to initialize the Skype Web SDK. You can follow the code below or use the **o365-skypeEnsureClient** code snippet.

        //ensures the skype client object is initialized
        var ensureClient = function() {
            var deferred = $q.defer();

            if (client != null)
                deferred.resolve();
            else {
                Skype.initialize({
                    apiKey: skypeSvc.config.apiKeyCC
                }, function (api) {
                    apiManager = api;
                    client = apiManager.UIApplicationInstance;
                    client.signInManager.state.changed(function (state) {
                        $rootScope.$broadcast("stateChanged", state);
                    });
                    deferred.resolve();
                }, function (er) {
                    deferred.resolve(er);
                });
            }

            return deferred.promise;
        };

7. The code in Step 6 ensures the Skype Web SDK is initialized, but you also need to ensure the user is signed in. Below ensureClient, add an **ensureSignIn** function to the **skypeSvc** object by using the **o365-skypeEnsureSignIn** code snippet. Notice that it checks the uses the **client.signInManager** to check the sign-in state and calls **signIn** if needed.

        //signs into skype
        skypeSvc.ensureSignIn = function() {
            var deferred = $q.defer();

            ensureClient().then(function() {
                //determine if the user is already signed in or not
                if (client.signInManager.state() == "SignedOut") {
                    client.signInManager.signIn(skypeSvc.config.initParams).then(function (z) {
                        //listen for status changes
                        client.personsAndGroupsManager.mePerson.status.changed(function (newStatus) {
                            console.log("logged in status: " + newStatus);
                        });

                        //In the future we can listen for new inbound conversations like this
                        client.conversationsManager.conversations.added(function (conversation) { });

                        //resolve the promise
                        deferred.resolve();
                    }, function (er) {
                        deferred.reject(er);
                    });
                }
                else {
                    //resolve the promise
                    deferred.resolve();
                }
            }, function(er) {
                deferred.reject(er);
            });

            return deferred.promise;
        }

8. Also notice that the code snippet above includes a line to listen for new inbound conversations. We won't fully implement that in this lab, so this is show just as a future reference.

        //In the future we can listen for new inbound conversations like this
        client.conversationsManager.conversations.added(function (conversation) { });

9. Next, you should create a **subscribeToStatus** function on the **skypeSvc** to check the status of user that is passed into it (queried using **client.personsAndGroupsManager.createPersonSearchQuery**). The function should also subscribe to status changes for the user. However, you don't want to subscribe the same user more than once, so keep track of subscriptions in a **userSubs** array. You can add this script block using the **o365-skypeSubscribeUser** code snippet.

        //subscribes to the status of a user
        var userSubs = [];
        skypeSvc.subscribeToStatus = function(id) {
            var deferred = $q.defer();

            //query for the user by their id
            var query = client.personsAndGroupsManager.createPersonSearchQuery();
            query.text(id);
            query.limit(1);
            query.getMore().then(function (items) {
                //ensure results came back
                if (items.length > 0)
                {
                    //assume the first match is the user
                    var person = items[0].result;
                    person.status.get().then(function (s) {
                        deferred.resolve(s);
                    });

                    //check if we have already subscribed to this user
                    var subMatch = null;
                    for (var i = 0; i < userSubs.length; i++) {
                        if (userSubs[i].id === id) {
                            subMatch = userSubs[i];
                            break;
                        }
                    }
                    if (!subMatch) {
                        //no subscription exists for this user, so create one
                        userSubs.push({ id: id, person: person });

                        //listen for status changes
                        person.status.changed(function(s) {
                            //broadcast the status change to listeners
                            $rootScope.$broadcast("statusChanged", { user: person, status: s });
                        });

                        //subscribe to the status changes
                        person.status.subscribe();
                    }
                }
                else
                    deferred.reject("No matches found");
            });

            return deferred.promise;
        };

10. Open the **controllers.js** file in the app folder and locate the **ticketsCtrl**. Update it dependency inject the **skypeSvc** you updated in the previous steps.

        .controller("ticketsCtrl", ["$scope", "helpdeskSvc", "skypeSvc", function($scope, helpdeskSvc, skypeSvc) {
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

11. Next, locate where **getTickets** returns tickets in the "then" promise. After setting $scope.tickets, you should ensure the user is signed into Skype for Business (using **skypeSvc.ensureSignIn**) and then loop through each ticket, subscribing to the ticket opener's status in Skype for Business (using  **skypeSvc.subscribeToStatus**).

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
                            $scope.$apply(); //this provides async ui update out of thread
                    });
                });
            });
        });

12. You also want to "listen" for status changes by Skype users. You already subscribed to these events with the Skype Web SDK in step 9, but you need to listen to the **statusChanged** event broadcast from the **skypeSvc**. You can do that as follows or using the **o365-skypeListenStatus** code snippet.

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

13. Finally, you need to modify the **view-tickets.html** file located in the **app** > **templates** folder to display the Skype presence of each user. Locate the the repeated table row and update the first table cell as follows or using the **o365-skypePresence** code snippet.

        <tbody>
            <tr ng-repeat="ticket in tickets">
                <td>
                    <span class="badge" ng-class="ticket.created_by.status"><span class="glyphicon glyphicon-minus"></span></span>
                    <span>{{ticket.created_by.name}}</span>
                </td>
                <td>{{ticket.title}}</td>
                <td>{{ticket.status}}</td>
            </tr>
        </tbody>

14. Launch the Skype for Business client by opening the Windows start menu and typing **Skype for Business 2016**. Once it launches, sign-in with the user account you have designated as the **Skype User**.

15. Open a browser and navigate to **http://localhost:8000**. After you sign-in with the account you designated as the **Browser User**, you should see a presence icon next to each ticket opener. It might take a few seconds, but the **TEST ACCOUNT** should light up with the correct presence from Skype for Business.

16. Try changing the presence of the Skype User in the Skype for Business 2016 client. After a few seconds, the new presence should display in the helpdesk app. 

	![Presence 1](http://i.imgur.com/YyT4wK5.png)

18. You have successfully integrated Skype for Business presence into an existing web application. In the next task you will integrate instant messaging, which the Skype Web SDK makes really easy with a conversation UI.

	![Presence 2](http://i.imgur.com/IXaVQOc.png)

<a name="Ex2Task3"></a>
#### Task 3 - Integrating Instant Messaging ####

In this task, you will continue to customize the Help Desk application to include instant messaging with ticket openers. The Skype for Business Web SDK include power controls that can help deliver a consistent Skype experience in your web applications.

1. Open the **skype.js** file created in the previous task and update it with a new **startConversation** function on the **skypeSvc** object that accepts a SIP address for a user an initiates a conversation with them using the Skype Web SDK. All you have to do to initiate the Skype conversation UI is to use the **apiManager** and the **renderConversation** function, providing a **DIV control** in the page that it can render in ("chatWindowInner" in the example below) and the conversation details (participants, modalities, etc). You can follow the code below or use the **o365-skypeStartConversation** code snippet.

        //start a conversation with a user
        skypeSvc.startConversation = function(sip) {
            //hide all containers
            var containers = document.getElementById("chatWindowInner").children;
            for (var i = 0; i < containers.length; i++) {
                containers[i].style.display = "none";
            }
            
            var chatSip = sip;
            var uris = [chatSip];
            var container = document.getElementById(chatSip);
            if (!container) {
                //this is a new conversation...create the window
                container = document.createElement("div");
                container.id = chatSip;
                document.getElementById("chatWindowInner").appendChild(container);
                var promise = apiManager.renderConversation(container, { modalities: ["Chat"], participants: uris });
            }
            else
                container.style.display = "block";
        };

2. Next, you need to update the **view-tickets.html** file in the **app** > **templates** folder to accommodate the conversation UI. Add the following HTML to the bottom of this file or use the **o365-skypeConversationUI** code snippet.

        <div id="chatWindow" ng-class="{'show': showChatWindow}">
            <span class="glyphicon glyphicon-remove close-chat" ng-click="closeChatWindow()"></span>
            <div id="chatWindowInner"></div>
        </div>

3. While you are in the **view-tickets.html** file, you should also update the presence indicator to have a click event to start a conversation. In AngularJS, click events are configured using **ng-click** attribute. Set **ng-click** to call **startChat(ticket)**.

		<span class="badge" ng-class="ticket.created_by.status" ng-click="startChat(ticket)"><span class="glyphicon glyphicon-minus"></span></span>

4. Next, open the **controllers.js** file in the **app** folder and locate the **ticketsCtrl**. At the bottom of this controller add a private **canChat** function that returns true/false based on the status and it's ability to accept instant messages. You can also add this using the **o365-skypeCanChat** code snippet.

        //helper function to check if a status can perform chat
        var canChat = function(status) {
            var chattableStatus = { 
                Online: true, Busy: true, Idle: true, IdleOnline: true, Away: true, BeRightBack: true,
                DoNotDisturb: false, Offline: false, Unknown: false, Hidden: false };
            return chattableStatus[status];
        };

5. Next, define a **startChat** function on the **$scope** object to initiate a conversation with the ticket opener. The function should accept a **ticket parameter** and check if the ticket opener is available to chat based on their availability (via the **canChat** function you just created). If the ticket opener is available for chat, you should call the **skypeSvc.startConversation** function from step 1 of this task. You can use the **o365-skypeStartChat** code snippet or follow the code below.
        
        //starts a chat
        $scope.startChat = function(ticket) {
            if (canChat(ticket.created_by.status)) {
                skypeSvc.startConversation("sip:" + ticket.created_by.email);
                $scope.showChatWindow = true;
            }
        };

6. Finally, add a **closeChatWindow** function on the $scope object to close the chat window. The **$scope.showChatWindow** property will dictate if chat window will be displayed. It is used in-conjunction with the **ng-show** attribute on the chat window markup.
        
        //closes the chat window
        $scope.closeChatWindow = function() {
            $scope.showChatWindow = false;
        };

7. Open a browser and navigate to **http://localhost:8000**. After you sign-in with the account you designated as the **Browser User**, you should see a presence icon next to each ticket opener. It might take a few seconds, but the **TEST ACCOUNT** should light up with the correct presence from Skype for Business. Click on the presence icon to initiate a conversation with that user. The Skype conversation UI should fly in from the right and allow you to have a instant messaging conversation with the user.

	![Integrated IM](http://i.imgur.com/i7MxUNC.png)

8. By successfully integrated presence and IM into the Help Desk application you have completed this exercise. Know that there are other Skype modalities in preview that you can develop against, including audio and video. Visit the Skype Quick Starts for more information.

<a name="Summary" />
## Summary ##

By completing this module, you should have:

- Manually registered an Office 365 Connector and used a webhook to send data into Office 365.
- Connected an existing web application to Office 365 via Office 365 Connectors, including the addition of "Connect to Office 365" buttons, handling callbacks from Office 365, and sending data into Office 365 via webhooks.
- Skype-enabled and existing web application using the Skype Web SDK.

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.
=======
<a name="HOLTop" />
# Integrating Conversations with Skype and Office 365 Connectors #

---

<a name="Overview" />
## Overview ##

Office 365 now enables developers to build a powerful new generation of **conversation**-driven applications. **Office 365 Connectors** and **Skype SDKs** are two developer platforms to lighting up conversations in your applications. Office 365 Connectors allow 3rd party applications to send data into Office 365 and Outlook in near real-time. Users can subscribe to relevant data/applications and have informative and actionable "cards" display in Outlook. Skype SDKs allow developers to being the power capabilities of Skype into their applications built for the web and mobile platforms.

In this module, you will explore development with Office 365 Connectors and Skype SDKs, including how to integrate them into existing web applications.

<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:

- Manually register an Office 365 Connector and use a webhook to send data into Office 365.
- Connect an existing web application to Office 365 via Office 365 Connectors, including the addition of "Connect to Office 365" buttons, handling callbacks from Office 365, and sending data into Office 365 via webhooks.
- Skype-enable and existing web application using the Skype Web SDK.

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater
- [Microsoft Office 2016][2]
- [Any Web Request Composer][2]

[1]: https://www.visualstudio.com/products/visual-studio-community-vs
[2]: https://portal.office.com
[3]: https://www.hurl.it

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [Building Office 365 Connectors](#Exercise1)
1. [Developing with the Skype Web SDK](#Exercise2)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise1"></a>
### Exercise 1: Building Office 365 Connectors ###

Office 365 Connectors are a great way to get useful information and content into your Office 365 Group. Any user can connect their group to services like Trello, Bing News, Twitter, etc., and get notified of the group's activity in that service. From tracking a team's progress in Trello, to following important hashtags in Twitter, Office 365 Connectors make it easier for an Office 365 group to stay in sync and get more done. Office 365 Connectors also provides a compelling extensibility solution for developers which we will explore in this Exercise.

In this Exercise, you will explore the developer options for working with Office 365 Connectors, including adding connectors, leveraging webhooks, and integrating Office 365 Connectors into existing web application via the "Connect to Office 365" button.

<a name="Ex1Task1"></a>
#### Task 1 - Getting Started with Office 365 Connectors ####

In this Task, you will explore Office 365 Groups and some of the existing Office 365 Connectors that are available.

1. Open a browser and navigate to [https://portal.office.com](https://portal.office.com "https://portal.office.com") and sign-in with the Office 365 credentials that were provided to you.

1. **Office 365 Connectors** for Groups are currently under developer preview. In order to access them, navigate to **OWA** using the URL [https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true](https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true "https://outlook.office.com/owa/#path=/mail&EnableConnectorDevPreview=true")

1. Once you are signed into OWA, locate the Office 365 Groups you are a member of in the lower left navigation.

	![Groups in OWA](Images/Mod4_Groups.png?raw=true "Groups in OWA")

    _Groups in OWA_

1. Create your own unique group by clicking the **+** (plus) sign to the right of the **Groups** title in the left navigation.

	![Add Group](Images/Mod4_AddGroup.png?raw=true "Add Group")

    _Add Group_

1. Provide a **name** and **description** for the new group and click **Create** (optionally add members once the groups has been created).

1. Select **More > Connectors** from the Group's top navigation (if More isn't an option, you might try navigating back to the group using the link in **Step 2**).

	![Connectors in group navigation](Images/Mod4_ConnectorsMenu.png?raw=true "Connectors in group navigation")

    _Connectors in group navigation_

1. Explore some of the Office 365 Connectors that are available out of the box.

	![OOTB Connectors](Images/Mod4_Connectors.png?raw=true "OOTB Connectors")

    _OOTB Connectors_

1. Locate the **Twitter** connector and click **Add**.

1. The Twitter connector requires you to sign-in with a Twitter account. To do this, click **Log in**.

	![Twitter Connector Log in](Images/Mod4_Twitter1.png?raw=true "Twitter Connector Log in")

    _Twitter Connector Log in_

1. After authorizing the Office 365 Connector for Twitter, you can select specific **users**/**hashtags** to follow and how **frequently** they show up in the Office 365 Group. Try to follow yourself or a hashtag.

	![Configure Twitter Connector](Images/Mod4_Twitter2.png?raw=true "Configure Twitter Connector")

    _Configure Twitter Connector_

1. Post a tweet that matches the criteria in Step 10 and see it show up in the Office 365 Group.

	![Twitter post via Office 365 Connector](Images/Mod4_Twitter3.png?raw=true "Twitter post via Office 365 Connector")

    _Twitter post via Office 365 Connector_

<a name="Ex1Task2"></a>
#### Task 2 - Leveraging Webhooks with Office 365 Connectors ####

Hopefully Task 1 helped to illustrate the power of Office 365 Connectors, but did little to showcase the unique developer opportunity. In this Task, you will explore Office 365 Connector webhooks and how developer can leverage then to send data into Office 365.

1. Navigate to the Office 365 Group you created in the previous Task and select **More > Connectors** from the Group's top navigation.

1. Locate the **Incoming Webhook** Connector and click **Add**.

	![Incoming Webhook](Images/Mod4_IncomingWebhook.png?raw=true "Incoming Webhook")

    _Incoming Webhook_

1. Specify a **name** for the incoming webhook (ex: Build 2016) and click **Create**.

1. The confirmation screen will display a **URL** that is the webhook end-point we will use later in this Task.

	![Webhook Confirmation](Images/Mod4_IncomingConfirmation.png?raw=true "Webhook Confirmation")

    _Webhook Confirmation_

1. Open a new browser tab and navigate to [https://www.hurl.it](https://www.hurl.it "https://www.hurl.it"), which is an in-browser web request composer similar to what Fiddler offers.

1. When the page loads, add the following details:
	- **Operation**: **POST**
	- **Destination Address**: **webhook URL** from **Step 4**
	- **Headers**: **Content-Type: application/json**
	- **Body**: **{ "text": "Hello from Build 2016" }**

	![Manual Webhook](Images/Mod4_ManualHook.png?raw=true "Manual Webhook")

    _Manual Webhook_

1. Accept the **Captcha** and click **Launch Request**. You should get a confirmation screen that looks similar to the following.

	![Webhook Manual Confirmation](Images/Mod4_ManualConfirm.png?raw=true "Webhook Manual Confirmation")

    _Webhook Manual Confirmation_

1. If you return to the Office 365 Group, you should be able to locate the message you sent into it via the webhook.

	![Message sent into Group via webhook](Images/Mod4_HookDone.png?raw=true "Message sent into Group via webhook")

    _Message sent into Group via webhook_

1. Although you sent a very simple message into the webhook, Office 365 Connectors support a much more complex message format. You can get more details on the message format by visiting [https://dev.outlook.com/Connectors/GetStarted](https://dev.outlook.com/Connectors/GetStarted "https://dev.outlook.com/Connectors/GetStarted").

<a name="Ex1Task3"></a>
#### Task 3 - Integrating "Connect to Office 365" into existing applications ####

Task 2 had you manually register a webhook for an Office 365 Connector. In this Task, you will modify an existing web application to register webhooks with Office 365 by leveraging a "Connect to Office 365" button. You will capture the webhook details in a custom callback and send messages to Office 365 when new records are created in the application.

This Task uses a starter project to serve as the existing application. The application is a Craigslist-style selling site named BillsList. You are tasked with enhancing BillsList to allow users to subscribe to listing categories and send messages to Office 365 Groups when new listings match the subscription criteria.

1. Open Windows Explorer and browse to the module's **Source\Ex1\Begin** folder.

1. Double-click the solution file (**BillsListASPNET.sln**) to open the solution in **Visual Studio Community 2015**.

1. The starter solution actually has two project...**BillsListASPNET** (the web application) and **BillsListASPNET.Data** (database project). Right-click the **BillsListASPNET.Data** project and select **Publish**.

1. On the **Publish Database** dialog, click **Edit** to configure the connection information.

	![Publish DB](Images/Mod4_PubDB1.png?raw=true "Publish DB")

    _Publish DB_

1. On the **Connection Properties** dialog enter **(localdb)\MSSQLLocalDB** for the **Server name** and click **OK**.

	![Connection Properties](Images/Mod4_DbCon2.png?raw=true "Connection Properties")

    _Connection Properties_

1. When you return to the **Publish Database** dialog, click **Publish** to publish the database to **LocalDb**.

1. When the database has finished publishing, press **F5** or start the debugger to test the starter project.

1. When the application loads, click **Listings** in the top navigation. This will prompt you to sign-in. Use the **Office 365 account** that was provided to you (it also supports Consumer/MSA accounts like outlook.com, live.com, hotmail.com, etc).

1. The **Listings** view is one of the views we want to modify to support subscriptions to Office 365 Connectors. Notice that listings also have a category link, which is the second view we will add the "**Connect to Office 365**" button.

	![Listings](Images/Mod4_Listings.png?raw=true "Listings")

    _Listings_

1. Close the browser to stop debugging and open the **Index.cshtml** file located in the web project at **Views\Items**.

1. Copy and paste the following markup at the end of the **H2** element (right after the text "**for sale**").

	````HTML
	<a style="float: right;" href="https://outlook.office.com/connectors/ConnectToO365?state=@Request.Url.AbsoluteUri&app_name=BillsList&app_logo_url=http://billslist.azurewebsites.net/images/logo_128.png&callback_url=https://localhost:44300/callback">
		<img src="https://connecttoo365.blob.core.windows.net/images/ConnectToO365Button.png" alt="Connect to Office 365" style="height: 32px;"></img>
	</a>  
	````

1. This snippet will add a "Connect to Office 365" button to the view. It passes the following parameters to Office 365:

	- **state**: optional state information...in our case we are passing the current view information so we can return to it after the connection has been established with Office 365
	- **app_name**: the name of the application (ex: BillsList)
	- **app_logo_url**: the application logo that will be displayed in Office 365 when messages are sent in via webhook
	- **callback**: the location that Office 365 will return webhook information to after the user has confirmed the connection (ex: https://localhost:44300/callback)

1. Next, open the **Category.cshtml** file located in the web project at **Views\Items**.

1. Copy and paste the following markup at the end of the **H2** element (right after "**@ViewData["category"]**").

	````HTML
	<a style="float: right;" href="https://outlook.office.com/connectors/ConnectToO365?state=@Request.Url.AbsoluteUri&app_name=BillsList%20(@ViewData["category"])&app_logo_url=http://billslist.azurewebsites.net/images/logo_128.png&callback_url=https://localhost:44300/callback">
		<img src="https://connecttoo365.blob.core.windows.net/images/ConnectToO365Button.png" alt="Connect to Office 365" style="height: 32px;"></img>
	</a>
	````

1. This markup snippet is slightly different. It sends a dynamic **app_name** to Office 365 that includes the category the user is subscribing to. This allows the user to subscribe to specific categories instead of ALL listings.

1. You might recall we are passing in a callback location of **https://localhost:44300/callback** to Office 365. However, the **Callback** controller does not yet exist...let's create it. Right click the **Controllers** folder in the web project and select **Add > Controller**.

1. Select **MVC Controller - Empty** for the controller type and name it **CallbackController**.

	![New Controller](Images/Mod4_CallbackCtrl.png?raw=true "New Controller")

    _New Controller_

1. Inside the **CallbackController** class, add the **o365-callbackctrl** code snippet by typing **o365-callbackctrl** and pressing **tab**.

	````C#
	// GET: Callback
	public ActionResult Index()
	{
		var error = Request["error"];
		var state = Request["state"];
		if (!String.IsNullOrEmpty(error))
		{
			 return RedirectToAction("Error", "Home", null);
		}
		else
		{
			 var group = Request["group_name"];
			 var webhook = Request["webhook_url"];
			 Subscription sub = new Subscription();
			 sub.GroupName = group;
			 sub.WebHookUri = webhook;

			 //set optional category
			 if (state.IndexOf("?c=") != -1)
				  sub.Category = state.Substring(state.IndexOf("?c=") + 3);

			 //save the subscription
			 using (BillsListEntities entities = new BillsListEntities())
			 {
				  entities.Subscriptions.Add(sub);
				  entities.SaveChanges();
				  return Redirect(state);
			 }
		}
	}
	````

1. This controller looks for information returned from Office 365 and saves it as a subscription. The specific information passed from Office 365 as parameters include:

	- **error**: error details if the connection with Office 365 failed (ex: user rejected the connection)
	- **state**: the state value that was passed in via the "Connect to Office 365" button. In our case, it could include the category the user subscribed to
	- **group_name**: the name of the group the user selected to connect to
	- **webhook_url**: the webhook end-point our application will use to send messages into Office 365

1. Almost done, just need to update the **Create** activity to send messages to the appropriate webhooks. Open the **ItemsController.cs** file located in the **Controllers** folder of the web project.

1. Towards the bottom of the class, add the **o365-callwebhook** code snippet by typing **o365-callwebhook** and pressing **tab** (resolve using dependencies if necessary).

	````C#
	private async Task callWebhook(string webhook, Item item)
	{
		var imgString = "https://billslist.azurewebsites.net/images/logo_40.png";
		if (Request.Files.Count > 0)
		{
			 //resize the image
			 Request.Files[0].InputStream.Position = 0;
			 Image img = Image.FromStream(Request.Files[0].InputStream);
			 var newImg = (Image)(new Bitmap(img, new Size(40, 40)));

			 //convert the stream
			 using (var stream = new System.IO.MemoryStream())
			 {
				  newImg.Save(stream, ImageFormat.Jpeg);
				  stream.Position = 0;
				  var bytes = new byte[stream.Length];
				  stream.Read(bytes, 0, bytes.Length);
				  imgString = "data:image/jpg;base64, " + Convert.ToBase64String(bytes);
			 }
		}

		//prepare the json payload
		var json = @"
			 {
				  'summary': 'A new listing was posted to BillsList',
				  'sections': [
						{
							 'activityTitle': 'New BillsList listing',
							 'activitySubtitle': '" + item.Title + @"',
							 'activityImage': '" + imgString + @"',
							 'facts': [
								  {
										'name': 'Category',
										'value': '" + item.Category + @"'
								  },
								  {
										'name': 'Price',
										'value': '$" + item.Price + @"'
								  },
								  {
										'name': 'Listed by',
										'value': '" + item.Owner + @"'
								  }
							 ]
						}
				  ],
				  'potentialAction': [
						{
							 '@context': 'http://schema.org',
							 '@type': 'ViewAction',
							 'name': 'View in BillsList',
							 'target': [
								  'https://localhost:44300/items/detail/" + item.Id + @"'
							 ]
						}
				  ]}";

		//prepare the http POST
		HttpClient client = new HttpClient();
		client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
		var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
		using (var response = await client.PostAsync(webhook, content))
		{
			 //TODO: check response.IsSuccessStatusCode
		}
	}
	````

1. This snippet takes the new listing details and sends it to Office 365 via **POST** to the webhook end-point.

1. Finally, locate the **Create** activity within the class. **Create** is overloaded, so select the one that is marked with **HttpPost** and has the **Item** parameter. Inside the using statement add the code below between **SaveChanges()** of the new listing and the **RedirectToAction()** statement. This identifies matching subscriptions and calls the appropriate webhooks.

	````C#
	//save the item to the database
	using (BillsListEntities entities = new BillsListEntities())
	{
		entities.Items.Add(item);
		var id = entities.SaveChanges();

		//loop through subscriptions and call webhooks for each
		var subs = entities.Subscriptions.Where(i => i.Category == null || i.Category == item.Category);
		foreach (var sub in subs)
		{
			 await callWebhook(sub.WebHookUri, item);
		}

		return RedirectToAction("Detail", new { id = item.Id });
	}
	````

1. It's time to test your work. Press **F5** or start the debugger to launch the application. When you click on **Listings** (and sign-in) the **Listings** view will have a "**Connect to Office 365**" button in the upper right. This button will also display on the **Category** view.

	![Connect to Office 365 button](Images/Mod4_ConnectTo.png?raw=true "Connect to Office 365 button")

    _Connect to Office 365 button_

1. View a specific **Category** and then click "**Connect to Office 365**". You should be redirected to a screen to select a Office 365 Group to connect to.

	![Select group to connect to](Images/Mod4_GroupSelect.png?raw=true "Select group to connect to")

    _Select group to connect to_

1. Select an Office 365 Group and click **Allow** to complete establish the connection with Office 365 and return to BillsList.

1. To test the connection, click **My Listings** and **Create listing** with the **category** you subscribed to. A Connector **Card** for the listing should almost immediately show up in the Office 365 Group.

	![Connector card](Images/Mod4_Card.png?raw=true "Connector card")

    _Connector card_

<a name="Exercise2"></a>
### Exercise 2: Developing with the Skype Web SDK ###

TODO

<a name="Ex2Task1"></a>
#### Task 1 - Setting up the code ####

TODO

<a name="Ex2Task2"></a>
#### Task 2 - Signing into Skype for Business ####

TODO

<a name="Ex2Task3"></a>
#### Task 3 - Testing the App ####

TODO

<a name="Summary" />
## Summary ##

By completing this module, you should have:

- Manually registered an Office 365 Connector and used a webhook to send data into Office 365.
- Connected an existing web application to Office 365 via Office 365 Connectors, including the addition of "Connect to Office 365" buttons, handling callbacks from Office 365, and sending data into Office 365 via webhooks.
- Skype-enabled and existing web application using the Skype Web SDK.
>>>>>>> dc76ed6b0d471156ede9be40ec115d699c4721d4
