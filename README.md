msaccess-gpg-demo
=================

GPG Usage Demo Application

Copyright © 2003 Brendan Kidwell

This Microsoft Access application demonstrates how to use the GNU Privacy Guard, available for free at http://www.gnupg.org/, to send encrypted database updates to a central FTP server. You must follow the installation instructions in this file before the program will work.

Portions of this program were copied The Access Web ( http://www.mvps.org/access/ ). In general, you may use the Visual Basic code found in this application however you wish, but be sure to read and respect any license information you may find at the top of each module.

Requirements
------------

* Microsoft Access 2000 or 2002 (XP)
* GNU Privacy Guard (Download the Windows version from the URL above.)
* access to an FTP server

Installation Overview
---------------------

Copy the two Access files, data.mdb and frontend.mdb, to a new folder ("c:\work\GPG Access Demo" for example).

Copy gpg.exe from the GPG folder into the folder where you put the database. Import someone's public key into a GPG keyring in the database's folder.

You must link the "addresses" table from data.mdb into frontend.mdb and then edit the module mdlSettings in frontend.mdb to reflect the FTP server name, the name on the GPG key, etc.

Importing the Gpg Key
---------------------

Export someone's public key to an ASCII file. (You should use your own personal key while you are experimenting.) For details on how to create a key pair and export to ASCII, see the GPG documentation.

Copy that file to the database's folder. Open a command prompt and go to that folder. Use GPG's import command:

```
c:\work\GPG Access Demo>gpg --homedir . --import KEY_FILE
```

where KEY_FILE is the name of the ASCII file you exported the public key to. GPG should respond by saying that 1 key was imported.

Next, set the trust on that key:

   c:\work\GPG Access Demo>gpg --homedir . --edit-key "YOUR_NAME"

where YOUR_NAME is the name on the key. GPG will give you its own command prompt. Type the command `trust` and select `5) I trust ultimately` and confirm with `yes`. Last, type `quit` to save your changes.

Configuring the Database
------------------------

Open frontend.mdb. From the File menu, choose Get External Data, then Link Tables. Locate and select the data file, data.mdb, and then choose to import the table `addresses`.

Verify that you have done this correctly by opening the "Browse / Edit Addresses" form from the Switchboard.

Next, go to the Modules section of the Database window and open `mdlSettings`. Fill in all the values according to your installation.

<table>
<tr><td>FTP_SERVER</td><td>the Internet name for the FTP server</td></tr>
<tr><td>FTP_PORT</td><td>FTP server's port (leave blank for default of 21)</td></tr>
<tr><td>FTP_USER</td><td>FTP user name</td></tr>
<tr><td>FTP_PASSWORD</td><td>FTP password</td></tr>
<tr><td>FTP_FOLDER</td><td>destination folder (e.g. "/home/addresses/updates"</td></tr>
<tr><td>GPG_RECIPIENT</td><td>the name on the key from the previous step</td></tr>
<tr><td>FILENAME_BASE</td><td>the beginning of the filename that is uploaded</td></tr>
</table>

If you give this database to more than one person, so they can each send updates to your FTP server, each one should have a different value for FILENAME_BASE, so you can tell who uploaded which file. The actual file name will also contain today's date.

Operation
---------

The basic operation of this system is fairly straightforward.

A remote user enters data into the database by way of forms in the front end --- Addresses is the only such form in this demo application. When the user wants to send his data to the FTP server, he uses the Upload command from the Switchboard. This command sends all of the data; there is no incremental update function in the demo.

At the other end, the owner of the key which was used in the encryption retrieves the update files from the FTP server and decrypts them. If he has many users sending him data, he can optionally collect all of the most recent updates"after decrypting them&rdquol;into a master database.

Notes
-----

The good stuff --- where the encryption and FTP uploading happens --- is in the form `Upload`. To view the Visual Basic code, go to the `Forms` section of the Database window and select `Upload` and then choose `Code` from the `View` menu.

The application makes use of the Microsoft Scripting Runtime library for filyststem access, because the built-in file handling functions in Visual Basic are somewhat limited and they have anachronistic QBasic-style syntax. You should already have this library because it should have been installed automatically with Windows. Even so, the application's reference to this library may get broken. If this happens, go into the Visual Basic editor by following the directions above and then choose References from the Tools menu. You will probably see a "Missing: Microsoft Scripting Runtime" in the list. Uncheck this and then find "Microsoft Scripting Runtime", and click OK.

It is important to note that the database is split across two files, data.mdb and frontend.mdb. The idea is to place all your data in a file by itself and all the rest of the program in another file (forms, reports, Visual Basic modules, etc.) Splitting applications this way is a common practice among Access programmers --- in fact, Access has a built-in command for splitting an existing database this way. This scheme makes it easy to work on the data as a distinct entity (compacting it, backing it up, encrypting and uploading it…) One apparent disadvantage is that the front end must be relinked to the data file when it is installed on a new machine.

If you want simply copy and paste from my sample application, here's what you need to do: Make sure your database is split the same way. Copy the form `Upload` and the modules `mdlFileStuff`, `mdlSettings`, `mdlShellWait`, and `mdlWhereAmI` to your front end. You may have to edit `Upload` to reflect the name of your data file.

The reason that I forced the application to use a copy of GPG installed in its own folder is to make it easy to distribute the finished database to a client; you wouldn't want to have to walk your client through installing GPG and your public key, and then telling the database where to find them. Simply give your client the two access files, `gpg.exe`, and the three files ending with `.gpg` and provide some mechanism or instructions for linking data.mdb into frontend.mdb.
