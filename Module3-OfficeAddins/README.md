<a name="HOLTop" />
# Building Office Add-ins with Web Development #

---

<a name="Overview" />
## Overview ##

Extending Office has been elusive opportunity for developers almost as long as Office has been around. With ~1.2 billion users, Office provides a powerful developer surface, especially when you consider the time typical information workers spend in Office. As Office has evolved to run in different form factors (PC, Mac, mobile, browser), Microsoft has improved the add-in development model to support these new platforms. The modern add-in is developed with standard web technologies such as HTML, CSS, and JavaScript. 

In this module, you will learn how to build Office add-ins using the web development tools and platforms you are familiar with both in and out of Visual Studio.

<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:

- Customize popular Office products such as Outlook, Excel, and Word
- Build Office Add-ins in Visual Studio 2015
- Convert existing web applications to host Office Add-ins
- Leverage open source tools to build Office Add-ins outside of Visual Studio
- Authenticate and connect to the Microsoft Graph from and Office Add-in

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater
- [Visual Studio Code][2]
- [Microsoft Office 2016][3]

[1]: https://www.visualstudio.com/products/visual-studio-community-vs
[2]: https://code.visualstudio.com/
[3]: https://portal.office.com

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [Converting Web Application to Mail Add-in](#Exercise1)
2. [Using Yeoman to Generate Office Add-in Projects](#Exercise2)
3. [Connecting to the Microsoft Graph from Add-in](#Exercise2)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise1"></a>
### Exercise 1: Converting Web Application to Mail Add-in ###

Mail Add-ins provide a powerful and contextual extension to Microsoft Outlook and Outlook Web Access (OWA). Contextual, because these add-ins can function different for each message or appointment they are launched against. These add-ins can be used in Read scenarios (for existing items in your inbox) and Compose scenarios (for new messages and appointments). In this Exercise, you will convert an existing ASP.NET MVC application to host a Mail Add-in.

The sample in Exercise 1 is a Mail CRM scenario that is popular for Mail Add-ins. The add-in will look up the email sender and allow you to view previous notes and capture new notes.

This exercise uses a starter project with basic project scaffolding pre-configured. The solution is built with ASP.NET MVC and uses LocalDB for data storage.

<a name="Ex1Task1"></a>
#### Task 1 - Convert the Web Application ####

In this Task, you will get the starter project up and running, then convert it to host and Office Add-in. Specifically, you will convert it to host a Mail Add-in activated when you read messages.

1. Open Windows Explorer and browse to the module's **Source\Ex1\Begin** folder.

2. The solution contains two projects. **MailCRM** is an ASP.NET MVC app that hosts all the web components of the solution. **MailCRM.Data** is a data project the defined the data model for the solution. You need to publish the database to **LocalDB**. Right-click the **MailCRM.Data** project and select **Publish**.

3.  On the **Publish Database** dialog, click on the **Edit** button to configure the connection information.

	![Publish database](Images/Mod3_PubDB1.png?raw=true "Publish database")
    
    _Publish database_

4. On the **Connection Properties** dialog enter **(localdb)\MSSQLLocalDB** for the **Server name** and click **OK**.

	![Database connection properties](Images/Mod3_PubDB2.png?raw=true "Database connection properties")
    
    _Database connection properties_

5. When you return to the **Publish Database** dialog, click the **Publish** button to publish the database to **LocalDb**.

6. With the database published, you should be able to test the web application. Press **F5** or start the debugger the launch the web application. Right now, the application should be blank, but this is an activity feed that will show real-time comments made on contacts.

	![Comment activity](Images/Mod3_ActivityEmpty.png?raw=true "Comment activity")
    
    _Comment activity_

7. After you stop debugging, right-click the **MailCRM** web project and select **Convert > Convert to App for Office Project...**

	![Convert project](Images/Mod3_Convert.png?raw=true "Convert project")
    
    _Convert project_

8. When the Office Add-in wizard launches, select **Mail** as the app type and click **Next**.

	![Wizard part1](Images/Mod3_Wizard1.png?raw=true "Wizard part1")
    
    _Wizard part1_

9. On the next wizard screen, select the option to have the add-in display for **Read form** on **Email messages** and click **Finish**.

	![Wizard part2](Images/Mod3_Wizard2.png?raw=true "Wizard part2")
    
    _Wizard part2_

10. The Office Add-in wizard will add a new **MailCRM.Office** project to the solution, but this project should contain only an XML manifest (**MailCRM.Office.xml**) file that defines how the add-in will appear in Outlook.

	![Solution with new Office add-in project](Images/Mod3_SlnWithOffice.png?raw=true "Solution with new Office add-in project")
    
    _Solution with new Office add-in project_

11. Open the **MailCRM.Office.xml** file located the **MailCRM.Office** project. Locate the **SourceLocation** element and update the value of the **DefaultValue** attribute to **~remoteAppUrl/Home/Agave**.

12. The **/Home/Agave** view we just referenced doesn't yet exist. Create it by right clicking the **Views > Home** folder (in the **MailCRM** project) and selecting **Add > View**.

	![Add view](Images/Mod3_AddView.png?raw=true "Add view")
    
    _Add view_

13. On the **Add View** dialog provide a **View name** of **Agave** and click the **Add** button.

	![Add agave view](Images/Mod3_AgaveView.png?raw=true "Add agave view")
    
    _Add agave view_

14. On the bottom of the new **Agave** view, insert the **o365-agaveview** code snippet using the key-combination **Ctrl-K**, **Ctrl-X**, and locate the **o365-agaveview** snippet under **My HTML Snippets**.

	>**NOTE:** the snippet added a script reference to **Office.js**. Office.js is the script file that provides the interop between web and Office. You will work more with this in **Task 2**. Also notice the **Office.initialize** reference. This is required to be called on any page that is displayed in an add-in. 

15. The conversion is almost complete, but the new **Agave** view doesn't have a matching activity in the **HomeController**. Open the **HomeContoller.cs** file in the **Controllers** folder of the **MailCRM** project. Add the following activity inside the **HomeController** class.

	````C#
	public ActionResult Agave()
	{
		return View();
	}
	````

16. The conversion is complete and this Task is complete. You could try debugging the add-in, but right now the add-in will display a blank page. You will add functionality in **Task 2**.

<a name="Ex1Task2"></a>
#### Task 2 - Leverage Office.js in Outlook ####

Office Add-ins can use just about any web technology (minus ActiveX). The become really powerful when combined with back-end services. In this Task, you will modify the add-in to leverage Office.js and communicate with some back-end REST services that are already hosted in the solution.

1. Open the **Agave.cshtml** file located at **Views > Home** in the **MailCRM** web project. 

2. Directly after the **$(document).ready(function () {** line add script that uses **Office.js** to get the sender of the email message (that the add-in is displayed for).

	````JavaScript
	Office.initialize = function (reason) {
		$(document).ready(function () {
			 //look up the sender
			 var sender = Office.context.mailbox.item.sender;
	````

3. Using the sender, you should perform a **REST** **GET** using the endpoint syntax /api/Comments?email=*{sender_email}*. Replace **sender_email** with email address of the sender, which can be found at **sender.emailAddress**. The **o365-httpget** code snippet can help you add a **JQuery** script block for performing an **HTTP** **GET**.

	````JavaScript
	$.ajax({
	url: "/api/Comments?email=" + sender.emailAddress,
		method: "GET",
		dataType: "json",
		success: function (data) {
		
		}
	});
	````

4. Next, add some additional script to loop through the data returned from the REST service and build the user interface in the **list** element. While you are at it, set the **header** element to the **sender.emailAddress** you retrieved from **Office.js** in **Step 2**.

	````JavaScript
	$("#header").html("Comments for " + sender.emailAddress);
	  $.ajax({
		url: "/api/Comments?email=" + sender.emailAddress,
			method: "GET",
			dataType: "json",
			success: function (data) {
				var html = "";
			$(data).each(function (i, e) {
				html += "<li class='list-group-item'>" + e.CommentText + "</li>";
			})
			$("#list").html(html);
			}
	  });
	````

5. Finally, add some script to **POST** new comments when the **submit** element is clicked. You can use the **o365-httppost** code snippet to generate a **JQuery** **HTTP** **POST** block and use the **/api/Comments** REST endpoint.

	````JavaScript
	$("#submit").click(function () {
		$.ajax({
				url: "/api/Comments",
					method: "POST",
				 contentType: "application/json; charset=utf-8",
				 dataType: "json",
				 data: JSON.stringify({ ContactEmail: sender.emailAddress, CommentText: $("#comment").val() }),
				 success: function (data) {
					var html = $("#list").html();
					  html = "<li class='list-group-item'>" + $("#comment").val() + "</li>" + html;
					  $("#list").html(html);
					  $("#comment").val("");
				 }
			});
	  });
	````

6. To test the add-in, press **F5** or start the debugger. If you are prompted to **Connect to Exchange email account**, provide the Office 365 credentials that were provided to you.

	![Connect to Exchange](Images/Mod3_debug.png?raw=true "Connect to Exchange")
    
    _Connect to Exchange_

7. Visual Studio will associate the add-in XML manifest with your Office 365 account and launch **OWA** in a browser to test the add-in. To test the add-in select a mail message and click on the **MailCRM** link below the recipients section of the reading pane.

	![Mail CRM](Images/Mod3_Mail1.png?raw=true "Mail CRM")
    
    _Mail CRM_

8. With the add-in expanded, you can add new comments for a contact by typing in the text box and clicking submit.

	![Mail Add-in](Images/Mod3_Mail2.png?raw=true "Mail Add-in")
    
    _Mail Add-in_

<a name="Ex1Task3"></a>
#### Task 3 - Leverage Add-in Command ####

In the final task for this exercise, you will update the add-in manifest to leverage an Add-in Command (aka - ribbon button) to launch the add-in. Add-in Commands are a great way to add your brand to Outlook as it can display an icon in the ribbon.

1. Open the add-in manifest file (**MailCRM.Office.xml**) in the **MailCRM.Office** project.'

2. Clear out everything in the manifest file and replace it with the **o365-addinmanifest** code snippet by using the key combination **Ctrl-K**, **Ctrl-X**, and selecting **o365-addinmanifest** from the **My XML Snippets**.

3. The code snippet added **VersionOverrides** element with the **Add-in Command** details. This element is ignored in older versions of Outlook. Take a moment to look at the details that are defined in this new section.

4. Test **Add-in Command** by press **F5** or start the debugger. Once **OWA** launches, open the **Outlook 2016** client. If prompted provide account details, provide the Office 365 account that was provided to you.

5. When **Outlook** launches, select a message and look in the ribbon for your **Add-in Command**.

	![Add-in command](Images/Mod3_AddinCommand.png?raw=true "Add-in command")
    
    _Add-in command_

6. Click the Add-in Command and the add-in will now display as a task pane, but otherwise functions the same.

	![Mail add-in as task pane](Images/Mod3_MailTaskPane.png?raw=true "Mail add-in as task pane")
    
    _Mail add-in as task pane_

<a name="Exercise2"></a>
### Exercise 2: Using Yeoman to Generate Office Add-in Projects ###

In Exercise 1, you leveraged Visual Studio to build an Office Add-in. Visual Studio provides the premier experience for development Office Add-ins. However, Microsoft offers a number of open source tools to develop Office Add-ins outside of Visual Studio.

In this exercise, you will use **Yeoman** to generate project scaffolding for an Office Add-in. You will also learn how to deploy and debug Office add-ins without Visual Studio.

<a name="Ex2Task1"></a>
#### Task 1 - Generate Project Scaffolding ####

In this task, you will use **Yeoman** to generate project scaffolding for an Office add-in.

1. Press the Windows start button and type Node.js, then locate and open the **Node.js command prompt** in the search results.

2. Change directories the C:\Project folder.

	````CMD
	cd C:\Projects
	````

3. Create a new directory for your Office Add-in named **ContactCleanup**

	````CMD
	mkdir ContactCleanup
	````

4. Change directories to the new directory.

	````CMD
	cd ContactCleanup
	````

5. Start the **Yeoman generator for Office Add-ins** by typing **yo office**.

	````CMD
	yo office
	````

6. The **Yeoman generator for Office Add-ins** is a command-line version of the new Office add-in wizard in Visual Studio. It will walk you though a few questions. Provide the following responses:

	- **Project name (display name)**: Contact Cleanup
	- **Office project type**: Task Pane Add-in
	- **Technology to use**: HTML, CSS & JavaScript
	- **Supported Office applications**: Excel

	![Yeoman generator for Office add-ins](Images/Mod3_Yeoman.png?raw=true "Yeoman generator for Office add-ins")
    
    _Yeoman generator for Office add-ins_

7. After providing the details above, Yeoman will use a number of tools to generate a starter project that will be used throughout this exercise. When it is complete, type the following to open the project in Visual Studio Code.

	````CMD
	Code .
	````

<a name="Ex2Task2"></a>
#### Task 2 - Explore Deploy/Debug Scenarios ####

Visual Studio handles the entire add-in hosting, deployment, and debugging with one simple button. For non-Visual Studio developers, it is important to understand how to perform these tasks. In Task 2 of Exercise 2, you will host the web application, deploy the add-in manifest, and test the add-in in Office.

1. Return to the command prompt you used in the previous Task. It should be opened to the project folder, but if no, change directories to that location.

2. Next, **gulp serve-static** in command window to allow a Gulp task to host the web application on port **8443**.

	````CMD
	gulp serve-static
	````

	![Gulp serve-static](Images/Mod3_Gulp.png?raw=true "Gulp serve-static")
    
    _Gulp serve-static_

3. Open a browser and navigate to **https://localhost:8443**. Yeoman not only created the basic project scaffolding for your add-in, it also generated and SSL certificate for hosting as add-ins are required the be hosted over SSL. 

	![Localhost](Images/Mod3_localhost.png?raw=true "Localhost")
    
    _Localhost_

4. If your browser displays certificate warning, click the option to **Continue to this site (not recommended)**, then view the certificate details and **manually install the self-signed certificate**.

	![Certificate Error](Images/Mod3_CertError.png?raw=true "Certificate Error")
    
    _Certificate Error_

	![Install Cert](Images/Mod3_InstallCert.png?raw=true "Install Cert")
    
    _Install Cert_

5. Now that the web application is host, you need to deploy the manifest so you can test the add-in. There are several ways to do this and we explore two of them. First, open a browser and navigate to [https://onedrive.live.com/](https://onedrive.live.com/ "https://onedrive.live.com/").

6. Sign-in with a **Microsoft Account (MSA)** such as **outlook.com**, **live.com**, **hotmail.com**, etc (same type of account you use to sign into Xbox). If you don't have one, you can get a free **OneDrive** account at [https://signup.live.com/](https://signup.live.com/ "https://signup.live.com/").

7. Once you are signed into **OneDrive**, select **New > Excel workbook** from the top navigation to launch a new workbook in **Excel Online**.

	![New Excel Workbook](Images/Mod3_NewExcel.png?raw=true "New Excel Workbook")
    
    _New Excel Workbook_

8. In **Excel Online**, select the **Insert** tab in the ribbon and then click the **Office Add-ins** button.

	![Insert Add-in](Images/Mod3_InsertAddin.png?raw=true "Insert Add-in")
    
    _Insert Add-in_

9. In the **Office Add-ins** dialog, locate and click the **Manage My Add-ins** link the upper right (near the **Refresh** link).

	![Manage Add-ins](Images/Mod3_ManageAddins.png?raw=true "Manage Add-ins")
    
    _Manage Add-ins_

10. Select **Upload My Add-in** from the selection menu and browse to the location of the **add-in XML manifest** file (should be in the root of your project folder).

	![Upload manifest](Images/Mod3_UploadManifest.png?raw=true "Upload manifest")
    
    _Upload manifest_

11. After uploading the **add-in XML manifest** file, the add-in should appear as a **task pane** in **Excel Online**.

	![Excel Online with Add-in](Images/Mod3_AddinEO.png?raw=true "Excel Online with Add-in")
    
    _Excel Online with Add-in_

12. You can now test and debug the add-in using the browser's developer tools (**F12**). However, this approach does not enable you to test the Office Add-in in the **Office client**. To do this, you can either use a special site collection in **SharePoint** (call an **App Catalog**) or use a **network file share**. We will use the network file share approach as it is easy and convenient for developers.

13. Open **Windows Explorer** and create a folder at the root the C: drive called **Apps** (C:\Apps).

14. Right-click the new folder and select **Properties** from the bottom of the menu. 

15. In the **Properties** dialog, select the **Sharing** tab and then click the **Share** button.

	![Properties](Images/Mod3_Share.png?raw=true "Properties")
    
    _Properties_

16. In the **File Sharing** dialog, click the **Share** button to create the share.

	![Share](Images/Mod3_Share2.png?raw=true "Share")
    
    _Share_

17. After configuring the share, the **Sharing** tab on the **Properties** dialog should display a **Network Path**. Copy or remember this path as we will configure Office to look at it for add-in manifests.

	![Network Path](Images/Mod3_Share3.png?raw=true "Network Path")
    
    _Network Path_

18. Next, open the **Excel 2016 client** and a new blank workbook.

19. Select **File > Options** to launch the **Excel Options** dialog.

20. Locate the **Trust Center** link in the left navigation and then click on the **Trust Center Settings** button to launch the **Trust Center** dialog.

21. In the **Trust Center** dialog, select **Trusted Add-in Catalogs** from the left navigation and add the network file share location from **Step 17** to the **Trusted Catalogs Table**. Make sure you check the **Show in Menu** checkbox.

	![Trusted catalogs](Images/Mod3_TrustedCats.png?raw=true "Trusted catalogs")
    
    _Trusted catalogs_

22. Close all the dialogs and return to the blank workbook in Excel. When you launch the Office Add-in dialog (via **Insert** tab > **My Add-ins**), you should now see a **Shared Folder** tab. This will load any add-in XML manifest that you place in the folder. Go ahead and try this with the manifest you built with Yeoman.

	![Shared Folder Insert Option](Images/Mod3_SharedFolderOption.png?raw=true "Shared Folder Insert Option")
    
    _Shared Folder Insert Option_

	>**NOTE:** A SharePoint App Catalog is a much more "Enterprise" approach for hosting add-ins. SharePoint makes it easier to manage add-in permissions and helps with the entire add-in lifecycle and license management. Network file shares can be more convenient for developers as was the case in this exercise. Regardless of catalog choise, Group Policy can be used to automatically configure Office for the enterprise (ultimately handling this entire task of the exercise).
 
<a name="Exercise3"></a>
### Exercise 3: Connecting to the Microsoft Graph from Add-in ###

Microsoft strategy around Office extensibility has focused on add-ins and APIs. So far, this module has focused completely on add-ins. In this last exercise you will combine the two by calling into the Microsoft Graph from your add-in.

<a name="Ex3Task1"></a>
#### Task 1 - Register the Application with Azure AD ####

In this task, you'll go through the steps of registering an app in Azure AD using the **Office 365 App Registration Tool**. You can also register apps in the **Azure Management Portal**, but the **Office 365 App Registration Tool** allows you to register applications without having access to the **Azure Management Portal**. Azure AD is the identity provider for Office 365, so any service/application that wants to use Office 365 data must be registered with it.

1. Open a browser and navigate to the **Office 365 App Registration Tool** at **[http://dev.office.com/app-registration](http://dev.office.com/app-registration "http://dev.office.com/app-registration")**.

2. The **Office 365 App Registration Tool** welcome screen will give you the option to use an existing Office 365 tenant or create a new Office 365 tenant. Click the **Sign in with your Office 365 account** and use your Office 365 account or the one that was provided to you.

	![App registration sign-in](Images/Mod3_AppRegSignin.png?raw=true "App registration sign-in")
    
    _App registration sign-in_

3. Once you are signed in, the app registration page will allow you to specify the details of your application. Use the details outlined below and then click the **Register App** button:

	- **App name**: **Contact Cleanup**
	- **App type**: **Web**
	- **Sign on URL**: **https://localhost:8443/app/auth/auth.html**
	- **Redirect URI**: **https://localhost:8443/app/auth/auth.html**
	- **App permissions**: **Contacts.ReadWrite**
 
4. Once the app registration is complete, the app registration page should display registration confirmation that includes a **Client ID** (below the **Register App** button). Copy this **Client ID** (or leave the browser up) for use in the next exercise. The registration tool will also list a **Client Secret**, but you can ignore that as we will be using an **Implicit OAuth Flow** that doesn't use the Secret.

	![Registration Confirmation](Images/Mod3_RegSuccess.png?raw=true "Registration Confirmation")
    
    _Registration Confirmation_

<a name="Ex3Task2"></a>
#### Task 2 - Leverage Dialog API for OAuth with Azure AD ####

In this task you will leverage a new Dialog API in your Office Add-in to perform OAuth with Azure AD. A dialog is necessary because an add-in can only display web content from domains that are pre-registered with the add-in. Federated sign-ins with Office 365 is common, and it would be impossible to add-in developers to anticipate all the federated sign-in domains. Also, some identity providers prevent their sign-in pages from being displayed in a frame (to prevent click-jacking exploits).

The Azure Active Directory Authentication Libraries (ADAL) are a set of libraries to simplify authentication with Azure AD. ADAL.js is the library you will use. Currently, it is only packaged with an AngularJS module, but you can ignore the AngularJS part.

1. Open your Office add-in project (created with Yeoman) in **Visual Studio Code**.

2. Open the **app.js** file located in the **app** folder.

3. After the variable **self** is defined, add properties to **self** for **clientId**, **tenant**, and **redirectUrl** with values from the app registration you performed in the previous task (**tenant** will be the domain of the Office 365 account you are using such as sometenant.onmicrosoft.com).

	````JavaScript
	self.clientId = "31b239c4-bd0e-4225-89e0-d59b7c52155b";
	self.tenant = "rzdemos.com";
	self.redirectUri = "https://localhost:8443/app/auth/auth.html";
	````

4. Next, open a command prompt to the location of your Office add-in project you started with Yeoman. If you need to stop the Gulp task, press Ctrl-C.

5. Install the **adal-angular** module by typing the command listed below in the command prompt:

	````CMD
	bower install adal-angular --save
	````

6. Return to **Visual Studio Code** and create a new folder named **auth** inside the **app** folder (ex: **app/auth**).

7. Next, create a new file named **auth.html** in the **auth** folder (ex: **app/auth/auth.html**).

8. **auth.html** will use **ADAL.js** to authenticate users of your app. Use the code snippet **o365-authjs** to complete this file by typing **o365-authjs** and pressing enter. The snippet will add script to check if the user is authenticated. If so, it will get an **access token** for the **Microsoft Graph** and return it to the parent (later in the exercise you will launch this page in a dialog). If the user isn't authenticated, it will force them through a sign-in and consent flow with Azure AD. Below is main block that handles that.

	````C#
	var user = authContext.getCachedUser();
				 
	  // Check if the user is cached
	  if (!user) {
		authContext.login();
	  } 
	  else {
		// Get access token for graph
		authContext.acquireToken("https://graph.microsoft.com", function (error, token) {
			// Check for success
				if (error || !token) {
					// Handle ADAL Error
					response.status = "error";
					  response.accessToken = null;
					  Office.context.ui.messageParent(JSON.stringify(response));
				 } 
				 else {
					// Return the roken to the parent
					  response.status = "success";
					  response.accessToken = token;
					  Office.context.ui.messageParent(JSON.stringify(response));
				 }
			});
	  }
	````

9. Next, let's modify the main user interface of the add-in. Open the **home.html** file located in the **app\home** folder and simplify the body to look like the following.

	````HTML
	<body>
		<div id="content-header">
			 <div class="padding">
				  <h1>Contact Clean-up</h1>
			 </div>
		</div>
		<div id="content-main">
			 <div class="padding">
			<h3>Status: <span id="status">Disconnected</span></h3>
				  <button id="btnSignin">Sign-in with Office 365</button>
				  <button id="btnSave" style="display: none;" disabled="true">Save Changes</button>
			 </div>
		</div>
	</body>
	````

10. Open the **home.js** file located in the **app\home** folder and modify the **(document).ready** script as seen below. This script loads the **auth.html** page in a dialog using a new **Dialog API** when the **btnSignin** element is clicked.

	````JavaScript
	var _dlg; // global pointer to the dialog

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function(reason){
	  $(document).ready(function(){
			app.initialize();

			$("#btnSignin").click(function () {
				 var url = "https://localhost:8443/app/auth/auth.html";
				 
				 Office.context.ui.displayDialogAsync(url, { height: 40, width: 40, requireHTTPS: true }, function (result) {
					  _dlg = result.value;
					  _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, dialogMessageReceived);
				 });
			});
	  });
	};
	````

11. _**dlg.addEventHandler** is adding an event handler to listen for messages sent from the dialog. Add a new **dialogMessageReceived** function with a **result** parameter to satisfy this handler.

	````JavaScript
	function dialogMessageReceived(result) {
		if (result && JSON.parse(result.message).status === "success") 
				$("#status").html("Connected");
		else
			$("#status").html("Authentication Failed");
	};
	````

12. You should be ready to test authentication now. If needed, start the web server by typing **gulp serve-static** in a command prompt.

	````CMD
	gulp serve-static
	````

13. Test the add-in the **Excel 2016** client. When you click on the Sign-in with Office 365 button, the add-in should launch a dialog for you to sign-in and consent to the app permissions.

	![Consenting app in dialog](Images/Mod3_Consent.png?raw=true "Consenting app in dialog")
    
    _Consenting app in dialog_

14. After signing in and consenting the application, the dialog should automatically close and the **Status** of the add-in should change to **Connected**.

	![Connected](Images/Mod3_Connected.png?raw=true "Connected")
    
    _Connected_

<a name="Ex3Task3"></a>
#### Task 3 - Calling the Microsoft Graph ####

In the final task, you will use the **access token** returned from the dialog and use it to call the **Microsoft Graph**. Specifically, you will query **Contacts** information stored in Office 365 and display it as a table in **Excel**.

1. Modify the **dialogMessageReceived** function to call into the Microsoft Graph. Notice the **Authorization** header we are setting on the request to "**Bearer** " plus the **access token** returned from the dialog. This is required to successfully get data back from the Microsoft Graph.

	````JavaScript
	function dialogMessageReceived(result) {
	  if (result && JSON.parse(result.message).status === "success") {
			$("#status").html("Connected");
			
			//close the dialog and call into Graph
			_dlg.close();
			var _token = JSON.parse(result.message);
			$.ajax({
				 url: "https://graph.microsoft.com/v1.0/me/contacts",
				 headers: {
					  "accept": "application/json",
					  "Authorization": "Bearer " + _token.accessToken
				 },
				 success: function (data) {
				//DO STUFF WITH DATA
			}
		});
	}
	else
		$("#status").html("Authentication Failed");
	};
	````

2. Inside the **success** callback, add script to loop through the returned contacts and write them to Excel using the **Office.context.document.setSelectedDataAsync** function.

	````JavaScript
	success: function (data) {
		var _officeTable = new Office.TableData();
			_officeTable.headers = ["First", "Last", "PrimaryEmail", "Company", "WorkPhone", "MobilePhone"];
					  
			$(data.value).each(function (i, e) {
				var _item = [
					e.givenName,
					  e.surname,
					  (e.emailAddresses.length > 0) ? e.emailAddresses[0].address : "",
					  (e.companyName) ? e.companyName : "",
					  (e.businessPhones.length > 0) ? e.businessPhones[0] : "",
					  (e.mobilePhone) ? e.mobilePhone : ""
			];

				_officeTable.rows.push(_item);
		});

		//toggle the buttons and set the data into Excel
			$("#btnSignin").hide();
			$("#btnSave").show();
			Office.context.document.setSelectedDataAsync(_officeTable, 
			{
				coercionType: Office.CoercionType.Table,
					  cellFormat: [{ 
					cells: Office.Table.All, 
					format: { width: "auto fit" } 
				}]
				 }, 
			function (asyncResult) {
				//TODO: Handle the result
			}
		);
	````

3. Test the add-in again in the **Excel 2016** client. After completing the sign-in dialog, the add-in should read **Connects** using the **Microsoft Graph** and write them to a table in Excel.

	![Table written to Excel](Images/Mod3_ExcelTable.png?raw=true "Table written to Excel")
    
    _Table written to Excel_

4. Office.js support a numerous powerful operations, including many product-specific. You could use these to deliver powerful scenarios in Office. Say we wanted to allow users to make bulk updates to contacts in Excel and then save them back to Office 365. You could use bindings and event handlers to know when the contact data changes in Excel. To see how you might start this, replace the **dialogMessageReceived** with the code snippet **o365-dialogMessageReceived** by typing **o365-dialogMessageReceived** and pressing **enter**.

5. Test the add-in one more time in the **Excel 2016** client. Once the add-in creates the table of Contacts, modify a contact and note that the add-in was notified of this change using a **BindingDataChanged** event handler.

	![Data changed notification](Images/Mod3_DataChanged.png?raw=true "Data changed notification")
    
    _Data changed notification_

6. Take a few moments to look at the script that created the binding and registered the event handler. This completes this Exercise and the Module.

<a name="Summary" />
## Summary ##

By completing this module, you should have:

- Customized popular Office products such as Outlook, Excel, and Word
- Built Office Add-ins in Visual Studio 2015
- Converted existing web applications to host Office Add-ins
- Leveraged open source tools to build Office Add-ins outside of Visual Studio
- Authenticated and connected to the Microsoft Graph from and Office Add-in