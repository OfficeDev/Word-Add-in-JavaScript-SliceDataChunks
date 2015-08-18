# Word Add-in: Send a Word document in chunks to a service

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
This sample shows how to use JavaScript in a Word 2013 task pane add-in to get the current document and slice it into chunks of data of user-defined sizes. The data could then be submitted to a service (such as an editing service, a translation service, or an e-book publishing service).

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Word 2013.
  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.
  

<a name="components"></a>
## Key components of the sample
The sample solution contains the following key files:

**WordDocumentEmitter project**

- WordDocumentEmitter.xml: The manifest file for the Word add-in.
- DocumentForEditing.docx: Starter document with 500 pages of text. 
 
**WordDocumentEmitterWeb project**

- App/Home/Home.html. The HTML user interface that is displayed in the task pane. 
- App/Home/Home.js. Logic that runs when the add-in is loaded. 


All other files are automatically provided by the Visual Studio project template for Office Add-ins.


<a name="codedescription"></a>
##Description of the code
The DocumentForEditing.docx file is set as the **StartAction** property of the task pane add-in. The document is large enough (500 pages) to be sliced into a number of discrete chunks of data. 

The sample demonstrates:

- How to use JavaScript to retrieve the selected value from a drop-down list.
- How to use the getFileAsync method to slice the file into chunks of data of varying sizes.
- How to retrieve the data from each slice of the file by using the getSliceAsync method.


<a name="build"></a>
## Build and debug ##

1. In Visual Studio, press F5 to build and deploy the sample add-in. The DocumentForEditing.docx file opens in Word.
2. In the task pane add-in, choose a size for the data chunk.
3. Click the **Publish now!** button. 

The add-in displays the number of slices and the size of each slice, along with buttons you can use to view the content of each slice.


<a name="troubleshooting"></a>
## Troubleshooting

- If the add-in starts with a blank document, ensure that the **StartAction** property of the WordDocumentEmitter project is set to *DocumentForEditing.docx* (not to *New Word Document*).
- If the add-in opens in read-only mode, click the **Enable editing** button.
- If the add-in does not appear in the task pane of the document, Choose **Insert > My Add-ins > Word Document Emitter**.


<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-BindContentControls).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Document.getFileAsync method](http://msdn.microsoft.com/library/office/apps/jj715284.aspx)
- [File.getSliceAsync method](http://msdn.microsoft.com/library/office/apps/jj715281.aspx)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
