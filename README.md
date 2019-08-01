# [ARCHIVED] Word Add-in: Send a Word document in chunks to a service

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary
This sample shows how to use JavaScript in a Word 2013 task pane add-in to get the current document and slice it into chunks of data in user-defined sizes. The data could then be submitted to a service (such as an editing service, a translation service, or an e-book publishing service).

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Word 2013.
  - Visual Studio 2013 (Update 5) or Visual Studio 2015, with Microsoft Office Developer Tools.  
  - Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.
  

<a name="components"></a>
## Key components of the sample
The sample solution contains the following key files:

**WordDocumentEmitter** project

- [WordDocumentEmitter.xml](https://github.com/OfficeDev/Word-Add-in-JavaScript-SliceDataChunks/blob/master/WordDocumentEmitter/WordDocumentEmitterManifest/WordDocumentEmitter.xml): The manifest file for the Word add-in.
- [DocumentForEditing.docx](https://github.com/OfficeDev/Word-Add-in-JavaScript-SliceDataChunks/blob/master/WordDocumentEmitter/DocumentForEditing.docx): Start Document with 500 pages of text. 
 
**WordDocumentEmitterWeb** project

- [App/Home/Home.html](https://github.com/OfficeDev/Word-Add-in-JavaScript-SliceDataChunks/blob/master/WordDocumentEmitterWeb/App/Home/Home.html). The HTML user interface that is displayed in the task pane. 
- [App/Home/Home.js](https://github.com/OfficeDev/Word-Add-in-JavaScript-SliceDataChunks/blob/master/WordDocumentEmitterWeb/App/Home/Home.js). Logic that runs when the add-in is loaded. 


<a name="codedescription"></a>
##Description of the code
The DocumentForEditing.docx file is set as the **Start Document** property of the task pane add-in. The document is large enough (500 pages) to be sliced into a number of discrete chunks of data. 

The sample demonstrates:

- How to use JavaScript to retrieve the selected value from a drop-down list.
- How to use the **getFileAsync** method to slice the file into chunks of data of particular sizes.
- How to retrieve the data from each slice of the file by using the **getSliceAsync** method.


<a name="build"></a>
## Build and debug ##

1. In Visual Studio, press F5 to build and deploy the sample add-in.
2. On the **Home** ribbon, click **Open** in the **Document Slicer** group.
2. In the task pane add-in, choose a size for the data chunk.
3. Click the **Publish now!** button. 

The add-in displays the number of slices and the size of each slice, along with buttons you can use to view the content of each slice.

>This sample displays the slice information to the user, but your add-in will probably send the data slices to a web service. The web service can then rebuild the presentation from the slices.


<a name="troubleshooting"></a>
## Troubleshooting

- If the add-in starts with a blank document, ensure that the **Start Document** property of the WordDocumentEmitter project is set to *DocumentForEditing.docx* (not to *New Word Document*).
- If the document opens in read-only mode, click the **Enable editing** button.
- If the add-in does not appear in the task pane of the document, Choose **Insert > My Add-ins > Word Document Emitter**.


<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-SliceDataChunks/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx) documentation on MSDN
- [Get the whole document from an add-in for PowerPoint or Word](https://msdn.microsoft.com/library/office/jj715279.aspx)
- [Document.getFileAsync method](http://msdn.microsoft.com/library/office/apps/jj715284.aspx)
- [File.getSliceAsync method](http://msdn.microsoft.com/library/office/apps/jj715281.aspx)
- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
