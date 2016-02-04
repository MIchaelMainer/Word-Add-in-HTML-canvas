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

To use the Image callouts Word ad-in sample, the following are required.

* [node.js](https://nodejs.org) to serve up the docx files.
* [npm](https://www.npmjs.com/) to install the dependencies.
* Word 2016 16.0.6326.0000 or higher, or any client that supports the Word Javascript API. This sample does a requirement check to see if it is running in a supported host.

> Note: Word for Mac 2016 does not support add-in commands at this time. This sample can run on the Mac without the add-in commands. 

TODO: test on a Mac

## Configure the add-in and Word

1. Unzip and run this [registry key](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/AddInCommandsUndark/EnableAppCmdXLWD.zip) to activate the add-in commands feature.
2. Install the Typescript definitions by running ```tsd install``` in the project's root directory at the command line.
3. Install the project dependencies by running ```npm install``` in the project's root directory. 
4.  



MUST have the reg key noted here : https://github.com/OfficeDev/Office-Add-in-Commands-Samples

TODO: note that add-in commands won't work on the Mac or iPad.
TODO: Create my own map.


tsd install - problem here is that the definitions are  out of date -- will cause issues. 
npm install
gulp copy:libs
gulp


For Mac (note that the add-in commands won't work -- update the manifest to point at the )
1.	Create a folder called “wef” in Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/
2.	Put the developer manifest in the wef folder (Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/wef)
3.	Open word application on Mac and click on inset->”my add-ins” drop down.


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

8) Install server.crt into the Trusted Root Certificate Authority store.




I put the xsd files in the VS xsd store. 