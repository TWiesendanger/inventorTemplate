# Inventor Template in C#

![InventorTemplate](https://user-images.githubusercontent.com/20424937/184002759-240fed33-e85f-407a-a262-a1f397967ea7.png)

- [Inventor Template in C#](#inventor-template-in-c)
  - [Introduction](#introduction)
  - [Planned features](#planned-features)
  - [Version](#version)
    - [Version higher than 2023](#version-higher-than-2023)
    - [Version lower than 2023](#version-lower-than-2023)
  - [Help providing for this template](#help-providing-for-this-template)
- [Using this template](#using-this-template)
  - [create new from template](#create-new-from-template)
  - [What to change after creation](#what-to-change-after-creation)
  - [Info dialog](#info-dialog)
  - [Use the installer to create a setup](#use-the-installer-to-create-a-setup)
  - [Build script](#build-script)
- [Troubleshooting](#troubleshooting)
  - [dotnet new not available in terminal](#dotnet-new-not-available-in-terminal)
  - [mt.exe missing](#mtexe-missing)

## Introduction

If you ever copied multiple things from other addins over and over again, then you will like this template. The idea is to provide a template that already includes alot of standard features that every addin needs. If you want to there is nothing to change at all. You could use a generated addin from this template right out of the box. Things like guids are already generated. There is also a build script to automatically deploy it to a predefined folder.

This template is based on the ***dotnet templating.*** See [here](https://docs.microsoft.com/en-us/dotnet/core/tools/custom-templates)

The following features are provided at the moment:

- Setup functioning Inventor addin in minutes
- default loggin with nLog
- Dark and Lighttheme support
- default buttons for all environments
- info dialog (show patch notes)
- Unload and loading of addin
- Reset of user interface
- loading settings file
- structured code base to extend
- Build script
- Inno setup installer script 
- documentation for using and extending the addin


## Planned features

- multi language support
- nuget package

Feel free to ask for other features by emailing me tobias.wiesendanger@gmail.com or by opening a issue.

## Version

The template is setup for Inventor 2023 but can be used for other versions. 

### Version higher than 2023

If you want to use this template for a version higher than 2023 you dont need to change anything if you dont want to.
Consider if you might want to reference a newer interop dll file.

To do this remove the existing one:

![image](https://user-images.githubusercontent.com/20424937/184006290-93fa7ad4-aa6a-4611-95f8-6a27300f7b00.png)

and reference the newer one that should be used. It can normaly be found here:  
`C:\Program Files\Autodesk\Inventor 2023\Bin\Public Assemblies`

Make sure to check the properties after. Embeding is set to true normaly and should be set to false.
Also copy should not be needed.

![image](https://user-images.githubusercontent.com/20424937/184006746-41b9489a-1bdd-4830-8efb-4ed28f50af2f.png)

### Version lower than 2023

The addin is setup to load with 2021 and should load with no problems. If you try to use it with an even lower version you need to change the button / icon methods because these are made for the light and dark theme that was added with 2021. Also make sure to use the correct interop dll as shown above.

## Help providing for this template

Feel free to fork and then create a pull merge request. The template itself cannot be run directly.
This is mainly because there are some variables that get only replaced when using `dotnet new`.

It's best to create an addin based on the template and develop inside of that. If it's working copy changes to the template.

# Using this template

Clone this project to an empty folder. After that run the following command to install it:

`dotnet new --install "C:\temp\inventorTemplate` in any terminal you like.
I prefer to use the [windows terminal](https://apps.microsoft.com/store/detail/windows-terminal/9N0DX20HK701?hl=de-de&gl=DE).

Make sure to replace the path with yours pointing to where the `.template.config` folder is.

You can test if the installation did work by typing:
`dotnet new --list`

![image](https://user-images.githubusercontent.com/20424937/184010640-ba1aaf9c-40bb-4a2f-a87a-a626583e8847.png)


## create new from template

To see what options are available you can always use this:
`dotnet new invAddin -h`

![image](https://user-images.githubusercontent.com/20424937/184011449-39b1308f-1a20-4ab1-8525-e77c4864e090.png)

As you can see there are currently two options that need to be provided.

This could look like this:
`dotnet new invAddin -n Testaddin -o "C:\temp\sampleAddin\Testaddin" -solFld "C:\temp\sampleAddin" -instFld "C:\ProgramData\Company\sampleAddin"`

Notice that there is also a -n option that defines the name of the addin and the an -o option which allows to determine where to place the project.
If you dont provide the -o option the current folder will be used.

> Use the addin name at the end of the -o path (see sample command)

> Make sure that the -solFld (or --solutionFolder) points to the -o path without the addin name.

> provide a installer path with the option -instFld (or --installFolder). This is where the addin will be deployed by the build script.


## What to change after creation

If you dont want to you dont need to change anything, but it makes sense to modify things like the description of the addin and a few other things.

Look for a folder called `Addin` and open a file called `*.addin`. Inside of it at least the description should be changed.

![image](https://user-images.githubusercontent.com/20424937/184017327-e15e396d-3285-4154-ba56-3dc748e52d3f.png)

If you plan to use the info dialog that is provided by default make sure to edit the file inside of the `Ressources` folder. [Info Dialog](##info-dialog)

If you want to use the installer you should have a look inside the `Installer` folder. [Installer](##info-dialog)

## Info dialog

There is a default dialog provided.

![image](https://user-images.githubusercontent.com/20424937/184018865-e5be260e-9ca1-4212-bc8a-3a97a92dd7bf.png)

If you plan to use it, you should edit the file inside of the `Ressources` folder. Whatever text is inside of the `versionhistory.txt` will be shown in this dialog.

![image](https://user-images.githubusercontent.com/20424937/184020251-848e4b8d-a3ba-4927-81de-064c3433e7b3.png)

## Use the installer to create a setup

The installer uses inno. Make sure to install [Setup](https://jrsoftware.org/isinfo.php).

Edit the sample script to your liking. At least the header needs some modifications.

```json
#define MyAppName "InventorTemplate"
#define MyAppVersion "0.0.0.1"
#define MyAppPublisher "Company"
#define MyAppURL "http://www.company.com"
```
## Build script

If you open the properties of the project and look for the tab `Build Events` there is a script that is run each time build is used.
There is at least one important step inside of it that needs to be done. 

```batch
"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\mt.exe" -manifest "SolutionFolder\InventorTemplate\Addin\InventorTemplate.manifest" -outputresource:"SolutionFolder\InventorTemplate\bin\Debug\InventorTemplate.dll";#2
```

This will embed the manifest inside the dll file. If this fails the addin will not load correctly. You can check if this was successfull by opening the dll file inside of visual studio.

![image](https://user-images.githubusercontent.com/20424937/184023722-1a770b7c-30fb-467b-8cfe-737e17e0513f.png)

If this is not present, something went wrong. Try checking the path to mt.exe. This can also be provided from other sources. If there is no mt.exe file see [Troubleshooting](##mt.exe-missing)

TODO Document the other steps and why they are there.
# Troubleshooting

## dotnet new not available in terminal

To use this you need to install the .NET SDK 3.1 or higher. Download it [here](https://dotnet.microsoft.com/en-us/download).

## mt.exe missing

If you dont have a mt.exe, then download the [Windows Dev Kit](https://developer.microsoft.com/de-de/windows/downloads/windows-10-sdk/).
It's sufficient to install `SDK for Desktop C++ x86 Apps`. After this look here: `C:\Program Files (x86)\Windows Kits\10\bin\<version>\<platform>`. 

How to change build script


How to add another button
How to provide Icons
What to change after creation
Implement button event
