Office JavaScript API 1.1.0.9
=============================
The Office JavaScript API includes objects, methods, properties, events, and enumerations that can be used from apps for Office to interact with Microsoft Office 2013 documents or mail items content.

This is a [bower](http://bower.io) package implementation of the Microsoft-authored NuGet package [Office JavaScript API 1.1.0.9](https://www.nuget.org/packages/Microsoft.Office.js/1.1.0.9).

Install
---
Install using bower

````
$ bower install microsoft.office.js [--save]
````

### Stuctural changes from version 1.1.0.7 to version 1.1.0.8
In version 1.1.0.7 the bower component created a folder named 'Microsoft.Office.js' in the bower components. This caused an issue when the scripts was directly embedded.

Since version 1.1.0.8 the folder name has changed to 'microsoft-office-js' to avoid this issue. 

Please make sure that the previous created folder was removed from the bower_components before / after upgrading to the new version.
