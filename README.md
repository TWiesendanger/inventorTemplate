# Inventor Template in C#

![InventorTemplate](https://user-images.githubusercontent.com/20424937/184002759-240fed33-e85f-407a-a262-a1f397967ea7.png)

- [Inventor Template in C#](#inventor-template-in-c)
  - [Introduction](#introduction)
  - [Planned features](#planned-features)
  - [Demo](#demo)
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
  - [Add a new button / command](#add-a-new-button--command)
    - [Create a Icon folder](#create-a-icon-folder)
    - [React on click event](#react-on-click-event)
    - [Add to interface](#add-to-interface)
  - [Globals](#globals)
  - [Add Settings](#add-settings)
- [Troubleshooting](#troubleshooting)
  - [dotnet new not available in terminal](#dotnet-new-not-available-in-terminal)
  - [mt.exe missing](#mtexe-missing)
  - [There was an error opening the file.](#there-was-an-error-opening-the-file)
  - [All my references are missing](#all-my-references-are-missing)
  - [Welcome to dll hell](#welcome-to-dll-hell)

## Introduction

If you ever copied multiple things from other addins over and over again, then you will like this template. The idea is to provide a template that already includes alot of standard features that every addin needs. If you want to there is nothing to change at all. You could use a generated addin from this template right out of the box. Things like guids are already generated. There is also a build script to automatically deploy it to a predefined folder.

This template is based on the ***dotnet templating.*** See [here](https://docs.microsoft.com/en-us/dotnet/core/tools/custom-templates)

The following features are provided at the moment:

- setup functioning Inventor addin in minutes
- default logging with nLog
- dark and Lighttheme support
- default buttons for all environments
- info dialog (show patch notes)
- unload and load of addin
- reset of user interface
- loading settings file
- structured code base to extend
- build script
- inno setup installer script 
- documentation for using and extending the addin

## Planned features

- multi language support
- nuget package

Feel free to ask for other features by emailing me tobias.wiesendanger@gmail.com or by opening a issue.

## Demo

![QuickStart](https://user-images.githubusercontent.com/20424937/184498033-71304b71-a91a-4568-893b-110edb9b116d.gif)

## Version

The template is setup for Inventor 2023 but can be used for other versions. 

### Version higher than 2023

If you want to use this template for a version higher than 2023 you dont need to change anything if you dont want to.
Consider if you might want to reference a newer interop dll file.

To do this remove the existing one:

![image](https://user-images.githubusercontent.com/20424937/184006290-93fa7ad4-aa6a-4611-95f8-6a27300f7b00.png)

and reference the newer one that should be used. It can normaly be found here:  
`C:\Program Files\Autodesk\Inventor 20xx\Bin\Public Assemblies`

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

![image](https://user-images.githubusercontent.com/20424937/184482872-a1bfd68f-bb32-4673-85f4-81af6b8ee955.png)

As you can see there is currently one option that needs to be provided.

This could look like this:
`dotnet new invAddin -n sampleAddin -o "C:\temp\sampleAddin\"  -instFld "C:\ProgramData\Company\sampleAddin"`

Notice that there is also a -n option that defines the name of the addin and an -o option which allows to determine where to place the project.
If you dont provide the -o option the current folder will be used.

> Use the addin name at the end of the -o path (see sample command)

> provide a installer path with the option -instFld (or --installFolder). This is where the addin will be deployed by the build script.

## What to change after creation

If you dont want to you dont need to change anything, but it makes sense to modify things like the description of the addin and a few other things.

Look for a folder called `Addin` and open a file called `*.addin`. Inside of it at least the description should be changed.

![image](https://user-images.githubusercontent.com/20424937/184017327-e15e396d-3285-4154-ba56-3dc748e52d3f.png)

If you plan to use the info dialog that is provided by default make sure to edit the file inside of the `Ressources` folder. [Info Dialog](##info-dialog)

If you want to use the installer you should have a look inside the `Installer` folder. [Installer](##Use-the-installer-to-create-a-setup)

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

Also make sure to change the license.txt file if you dont plan to release your addin with a GNU license.
## Build script

If you open the properties of the project and look for the tab `Build Events` there is a script that is run each time build is used.
There is at least one important step inside of it that needs to be done.

```batch
SET PATHTORUNIN=%CD:bin\Debug=%
"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\mt.exe" -manifest "%PATHTORUNIN%Addin\InventorTemplate.manifest" -outputresource:"%PATHTORUNIN%bin\Debug"
```

This will embed the manifest inside the dll file. If this fails the addin will not load correctly. You can check if this was successfull by opening the dll file inside of visual studio.

![image](https://user-images.githubusercontent.com/20424937/184023722-1a770b7c-30fb-467b-8cfe-737e17e0513f.png)

If this is not present, something went wrong. Try checking the path to mt.exe. This can also be provided from other sources. If there is no mt.exe file see [Troubleshooting](##mt.exe-missing)

After this step, if the addin is already existing, the addin file will be deleted. After this the file has to be copied from your debug folder to this folder.
This will be done in the next step that looks like this:

```batch
echo - Copy the .addin file folder into the standard Inventor Addins folder.
XCopy "%PATHTORUNIN%Addin\InventorTemplate.Inventor.addin" "C:\ProgramData\Autodesk\Inventor Addins" /y
```

The last and most important stop is to copy everything belonging to the addin into the predefined installfolder. Here the path that you provided at the begining is used.

## Add a new button / command

There are some default buttons that show how a buttons has to be setup, but I would like to show a bit more insight here, because that is a step that probably everyone needs to do.
Also every addin out there does that a bit different.

Start in StandardAddinServer.cs by adding a new field. There should be two field already created at Line 34 and 35.

![image](https://user-images.githubusercontent.com/20424937/184493263-147271c9-d7cd-4ab9-9479-0aaf5318cac3.png)

Remove the existing ones if you dont need them and add new ones here. If you want to follow the microsoft best practise use a "_" before the name.

Next, look inside of the `Activate` method. There is a sample line that looks like this:

```csharp
_info = UiDefinitionHelper.CreateButton("Info", "InventorTemplateInfo", @"UI\ButtonResources\Info", theme);
```

Follow the exact same principe:

```csharp
_yourFieldName = UiDefinitionHelper.CreateButton("DisplayName", "InternalName", @"UI\ButtonResources\YourCommandName", theme);
```
For InternalName it might be a goode idea to also add your addin name as a prefix to not risk any conflicting names. This has to be a unique name and also does not collide with already existing ones from autodesk inventor.
Make sure to also add it to `_buttonDefinitions`. This makes sure that the button can be deleted when unloading the addin.

### Create a Icon folder

To provide icons for your own button, make sure to create a new Folder inside of `UI\ButtonRessources` Name it the as the name you used for creating the button.
Inside of it you need to provide 4 different icons. There are two sizes that need to be provided. (16x16 adn 32x32) Both of them have to be named exactly like those in the default folders.

Make sure that you provide icons for the dark and the lightTheme. If you dont want to support the darkTheme, feel free to just copy lightTheme icons and rename them.
They have to be provided as png.

### React on click event

To react on the user clicking your button, you need to add a case in the `UI\Button.cs`

```csharp
case "InventorTemplateDefaultButton":
    MessageBox.Show(@"Default message.", @"Default title");
    return;
case "InventorTemplateInfo":
    var infoDlg = new FrmInfo();
    infoDlg.ShowDialog(new WindowWrapper((IntPtr)Globals.InvApp.MainFrameHWND));
    return;
default:
    return;
```

Just add another case exactly the same way as those that are already provided. The case string has to be the internal name you defined while creating the button.
After that call your method from here, to do whatever the button should do.

### Add to interface

To make sure that your button is shown in the interface you need to edit the `AddToUserInterface`. Depending on how and where you want to show it, the steps will be different.

IF you want to create a new **tab** then do the following:

```csharp
var yourTab = UiDefinitionHelper.SetupTab("DisplayName", "InternalName", onWhatRibbonToPlace); // use the alreay provided variables for onWhatRibbonToPlace
```
This is a tab

![image](https://user-images.githubusercontent.com/20424937/184494993-e30c9f9a-089e-413e-8f95-22e33e231f39.png)

Make sure to also add the tab to `_ribbonTabs`. As with the buttondefinitions, this allows to remove them easily.

Add a panel to the tab:

```csharp
var yourPanel = UiDefinitionHelper.SetupPanel("DisplayName", "InternalName", yourTab);
```

This is a panel

![image](https://user-images.githubusercontent.com/20424937/184495425-6039a5e1-1bf0-4d78-b0e0-f01375fbfcaf.png)

Make sure to also add the tab to `_ribbonPanels`. As with the buttondefinitions, this allows to remove them easily.

As a last step, check if your button is not null (could be if Buttodefinition.Add did not work), and then add it to your panel.

```csharp
if (_yourButton_ != null)
{
    var yourButtonRibbon = yourPanel.CommandControls.AddButton(_yourButton_, true);
}
```

## Globals

There is a file called `Globals.cs`, which can be used to get a reference to the `Application`. Use it like this:

```csharp
Globals.InvApp
```
## Add Settings

There is one sample setting already provided and another one is used as the loggin path.

Look inside of the `appsettings.json`. There are currently two sections. One called Logging and one called AppSettings. To get a reference to those, use this:

```csharp
var logSettings = config.GetSection("Logging").Get<AppsettingsBinder>();
```

`AppsettingsBinder.cs` is used to allow a mapping to properties. If you create a new section inside of `appsettings.json` make sure to also create a new class inside of `AppsettingsBinder.cs`

For each new setting that you create make sure to add a property to the class. This will later allow to access them with intelisense.

# Troubleshooting

## dotnet new not available in terminal

To use this you need to install the .NET SDK 3.1 or higher. Download it [here](https://dotnet.microsoft.com/en-us/download).

## mt.exe missing

~~If you dont have a mt.exe, then download the [Windows Dev Kit](https://developer.microsoft.com/de-de/windows/downloads/windows-10-sdk/).
It's sufficient to install `SDK for Desktop C++ x86 Apps`. After this look here: `C:\Program Files (x86)\Windows Kits\10\bin\<version>\<platform>`.~~

There is now an `mt.exe` inside of the ressource folder that gets referenced by the build script. So this should no longer be a problem.

## There was an error opening the file.

If your addin for some reason doesn't load, check what the Add-Ins Dialog shows.

![image](https://user-images.githubusercontent.com/20424937/184718727-d089899e-7517-428a-92e3-d5d2d972936d.png)

If there was an error opening the file is shown. Make sure that it points to the right dll file. It's not enough to just give the path.
So this will not work:

![image](https://user-images.githubusercontent.com/20424937/184719392-fac1d54d-90f1-4e83-abd2-820ba044cc01.png)

but this is fine:

![image](https://user-images.githubusercontent.com/20424937/184719423-39b14420-eca1-409b-9305-bda6ab16b123.png)

## All my references are missing

Try dotnet restore inside of the developer console to download any referenced nuget packages.
For all the other references, there are different things that can lead to this:

1. Try to unload the project and reload with dependencies.
2. Remove project and readd it to the solution.
3. Remove this from the project file (make sure to get a backup from the file first)

```csharp
<Target Name="EnsureNuGetPackageBuildImports" BeforeTargets="PrepareForBuild">
  <PropertyGroup>
    <ErrorText>This project references NuGet package(s) that are missing on this computer. Enable NuGet Package Restore to download them.  For more information, see http://go.microsoft.com/fwlink/?LinkID=322105. The missing file is {0}.</ErrorText>
  </PropertyGroup>
  <Error Condition="!Exists('$(SolutionDir)\.nuget\NuGet.targets')" Text="$([System.String]::Format('$(ErrorText)', '$(SolutionDir)\.nuget\NuGet.targets'))" />
</Target>
```

This should allow to fix everything other than nuget packages. To restore those use this:
`dotnet restore` inside of the developer shell.

## Welcome to dll hell

There are references to dll files which could lead to problems if another addin or even inventor itself is already using this in another version. This should be solved by using the `CurrentDomain_AssemblyResolve` method. The addin deploys everything to it's installfolder and those dll files should be loaded.