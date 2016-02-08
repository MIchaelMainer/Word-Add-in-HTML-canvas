Summary: Learn how you can use Typescript and the HTML Canvas API to extend the image editing capability in a Word document.
Name: Word-Add-in-Typescript-HTML-Canvas

# Image callouts Word add-in sample: load, edit, and insert images 

This Word add-in sample shows you how to:

1. Create a Word add-in with Typescript.
2. Load images from Word into the add-in.
3. Edit images in the add-in by using the HTML canvas API and insert the images into a Word document.
4. Implement add-in commands that both launch an add-in from the ribbon and run a script from a context menu.
5. Use the Office UI Fabric to create a seamless Word user experience.

TODO: Add GIF that shows the sample running. 

## Prerequisites

To use the Image callouts Word add-in sample, the following are required.

* [node.js](https://nodejs.org) to serve up the docx files.
* [npm](https://www.npmjs.com/) to install the dependencies.
* Word 2016 16.0.6326.0000 or higher, or any client that supports the Word Javascript API. This sample does a requirement check to see if it is running in a supported host for the JavaScript APIs. 
* Clone this repo to your local computer.

> Note: Word for Mac 2016 does not support add-in commands at this time. This sample can run on the Mac without the add-in commands. 

Since add-in commands that run scripts don't provide a UI to accept invalid certificates, you'll need to either run a proxy that provides its on certificate (like Fiddler), or setup your own certificates. If you want to use your own certificates, then OpenSSL is required.  

## Create developer certificates

You'll probably want to create your own certificates to run this sample on your development computer.

### Setup on a Windows computer

1. Follow the instructions on [Didier Steven's blog](http://blog.didierstevens.com/2015/03/30/howto-make-your-own-cert-with-openssl-on-windows/) for creating a certificate authority and server certificate. We suggest that you give the certificate authority certificate a *Common Name* of *localhost-ca*. The server certificate must have a *Common Name* of *localhost*.
2. Move the certificates you created to the root of this project.
3. Update gulpfile.config.json with the passphrase for the certificate. 

### Setup on a Mac

TODO

## Configure the add-in and Word

1. Unzip and run this [registry key](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/AddInCommandsUndark/EnableAppCmdXLWD.zip) to activate the add-in commands feature. This is required while add-in commands are a **preview feature**.
2. Install the Typescript definitions identified in tsd.json by running ```tsd install``` in the project's root directory at the command line. Note that the TypeScript definitions are out of date and will cause errors. You'll need to fake the missing definitions until the official definitions are updated on DefinatelyTyped. The definitions are in a directory called typings.
3. Install the project dependencies identified in package.json by running ```npm install``` in the project's root directory. 
4. Copy the Fabric and JQuery files by running ```gulp copy:libs```.
5. Add-in commands require HTTPS so you'll need to create a local certificate authority cert, and a server cert and key.  Place the files server.key, server.crt, and ca.crt at the root of this application. Alternatively, you can run this sample using a proxy like Fiddler that supplies its own certificate. 
6. Run the default gulp task by running ```gulp``` from the project's root directory. If the TypeScript definitions aren't updated, you'll get an error here. 
7. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx) and place the [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) manifest file in it.

You've deployed this sample add-in at this point. Now you need to let Word know where to find the add-in.

### Word 2016 for Windows setup

1. Launch Word and open a document.
2. Choose the **File** tab, and then choose **Options**.
3. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
4. Choose **Trusted Add-ins Catalogs**.
5. In the **Catalog Url** box, enter the network path to the folder share that contains manifest-word-add-in-canvas.xml and then choose **Add Catalog**.
6. Select the **Show in Menu** check box, and then choose **OK**.
7. A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Word. 

### Word 2016 for Mac setup

## Run the add-in in Word 2016 for Windows

1. Open a Word document. 
2. On the **Insert** tab in Word 2016, choose **My Add-ins**. 
3. Select the **Shared folder** tab.
4. Choose **Image callout add-in**, and then select **Insert**.
5. If add-in commands are suported by your version of Word, the UI will inform you that the add-in was loaded. You can use  the Developer tab to load the add-in in the UI and to insert an image into the document. You can also use the right-click context menu to insert an image into the document. 
6. If add-in commands are not supported by your version of Word, the add-in will load in a task pane. You'll need to insert a picture into the Word document to use the functionality of the add-in.
7. Select an image in the Word document, and load it into the taskpane by selecting *Load image from doc*. You can now insert callouts into the image. Select *Insert image into doc* to place the updated image into the Word doc. The add-in wil generate placeholder descriptions for each of the callouts. 

## FAQ

* Will add-in commands work on Mac and iPad? No, they won't work on the Mac or iPad as of the publication of this readme.
* Why doesn't my add-in show up in the **My Add-ins** window? Your add-in manifest may have an error. I suggest that you validate the manifest against the [manifest schema](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD).
* Why doesn't the function file get called for my add-in commands? Add-ins commands require HTTPS. Since the add-in commands require TLS, and there isn't a UI, you can't see whether there is a certificate issue. If you have to accept an invalid certificate in the taskpane, then the add-in command will not work.  

## Questions and comments

We'd love to get your feedback about the Image callout Word add-in sample. You can send your questions and suggestions to us in the [TODO](TODO) section of this repository.

Questions about add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [office-js], [word-addins], and [API]. We are watching these tags.

## Learn more

Here are more resources to help you create Word Javascript API based add-ins:

* [Office Add-ins platform overview](https://msdn.microsoft.com/EN-US/library/office/jj220082.aspx)
* [Word add-ins](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)
* [SillyStories sample](https://github.com/OfficeDev/Word-Add-in-SillyStories) - learn how to load docx files from a service and insert the files into an open Word document.

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.




STUFF TO ADDRESS

For Mac (note that the add-in commands won't work -- update the manifest to point at the )
1.	Create a folder called “wef” in Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/
2.	Put the developer manifest in the wef folder (Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/wef)
3.	Open word application on Mac and click on insert->”my add-ins” drop down.

Add-in commands require HTTPS 
The gulp-connect server has expired certificates so to run this yourself, you'll need to either: generate your own certificates, or run a proxy like Fiddler that provides certificates. 

OpenSSL

Install OpenSSL if you don't already have it.
https://wiki.openssl.org/index.php/Binaries

http://blog.didierstevens.com/2015/03/30/howto-make-your-own-cert-with-openssl-on-windows/

1) Open cmd window, Set these environment variables
set RANDFILE=c:\demo\.rnd
set OPENSSL_CONF=C:\OpenSSL-Win32\bin\openssl.cfg

2) Start OpenSSL by typing 
    ```c:\OpenSSL-Win32\bin\openssl.exe```

3) Create RSA key for the root CA and store it in ca.key:
    ```genrsa -out ca.key 4096```

4) create our self-signed root CA certificate ca.crt; you’ll need to provide an identity for your root CA:
    ```req -new -x509 -days 1826 -key ca.key -out ca.crt```
    
5) create our subordinate CA that will be used for the actual signing. First, generate the key:

    ```genrsa -out server.key 4096```
    
6) Then, request a certificate for this subordinate CA:

    ``` req -new -key server.key -out server.csr```
    
7) process the request for the subordinate CA certificate and get it signed by the root CA.

    ``` x509 -req -days 730 -in server.csr -CA ca.crt -CAkey ca.key -set_serial 01 -out server.crt```

7.5) To use this subordinate CA key for Authenticode signatures with Microsoft’s signtool, you’ll have to package the keys and certs in a PKCS12 file:    

    ```pkcs12 -export -out server.p12 -inkey server.key -in server.crt -chain -CAfile ca.crt```

8) Install server.crt into the Trusted Root Certificate Authority store. (Confirm this).
