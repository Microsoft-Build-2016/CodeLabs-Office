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

TODO

<a name="Ex2Task1"></a>
#### Task 1 - Setup the Code ####

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

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.
