# Sharepoint
Learning SharePoint.

## Sharepoint Basics
https://support.microsoft.com/en-us/office/get-started-with-sharepoint-909ec2f0-05c8-4e92-8ad3-3f8b0b6cf261

<img width="400" alt="image" src="screenshots/sharepoint-link.png">

Make sure you're logged in using your Microsoft 365 account to use SharePoint.

### Microsoft Forms
2 ways to create forms:
- [Microsoft Forms](https://support.microsoft.com/en-us/office/create-a-form-with-microsoft-forms-4ffb64cc-7d5d-402f-b82e-b1d49418fd9d)
- [Microsoft Forms web part on a SharePoint site](https://support.microsoft.com/en-us/office/use-the-microsoft-forms-web-part-on-a-sharepoint-site-d4b4d3ce-7860-41e4-8a98-76380efe7256) (Shown below)

#### Create a form using Microsoft Forms web part on a SharePoint site
1. Go to your SharePoint site. [Site creation instructions here.](#create-a-new-site-collection)
2. Click Edit.

   <img width="1200" alt="image" src="screenshots/mf-site-edit.png">
3. Click the plus icon to add a new web part.

   <img width="600" alt="image" src="screenshots/mf-add-new-web-part.png">
4. Search for Microsoft Forms and select it.

   <img width="500" alt="image" src="screenshots/mf-add-forms.png">
5. Name the form and hit Create.

   <img width="350" alt="image" src="screenshots/mf-name-form.png">
6. Create the form.
  
   <img width="900" alt="image" src="screenshots/mf-form-created.png">
7. When you're done creating your form, go back to your SharePoint in Microsoft 365 page. 
   Make sure Collect responses is selected, then select OK to refresh so you're seeing the most updated content.
  
   <img width="250" alt="image" src="screenshots/mf-collect-responses-ok.png">
8. (Optional) Add the form to the navigation menu of your SharePoint site if you'd like. [Reference](https://youtu.be/snoD8iXWcEc?si=UY-kEEhv0hTitt9G).

   <img width="500" alt="image" src="screenshots/mf-add-forms-to-nav.png">
   
#### Show results
2 ways (In Forms or Sharepoint).
1. In Forms: Site > Edit > Select form > Edit Properties > Visit Form web address > View responses.
2. In Sharepoint: Site > Edit > Select form > Edit Properties > Select "Show form results" > Ok > Republish.

   <p>
     <img width="250" alt="image" src="screenshots/mf-view-form-results.png">
   &nbsp;
     <img width="600" alt="image" src="screenshots/mf-form-results.png">
   </p>

#### Delete the form: 
1. SharePoint site > Edit > Select form > Delete web part > Republish
2. Go to [Microsoft Forms](https://forms.office.com/Pages/DesignPageV2.aspx)
   - Select the form

     <img width="550" alt="image" src="screenshots/mf-select-form.png">
   - Click form name from the top center > Select group

     <img width="300" alt="image" src="screenshots/mf-select-group.png">
   - Click the three dots on the form > Delete
   
     <img width="600" alt="image" src="screenshots/mf-delete-group-forms.png">
    
## Sharepoint Development
https://learn.microsoft.com/en-us/sharepoint/dev/

### Overview
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview


# SharePoint Framework (SPFx)
## Setup your Microsoft 365 tenant for development
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant

### Create app catalog
- I got my Microsoft 365 tenant through the ["365 developer" program](https://developer.microsoft.com/en-us/microsoft-365/dev-program). But that program is not available anymore.
- You need an app catalog to upload and deploy web parts.
- Since my email is `<myname>@lns4.onmicrosoft.com` so my tenant prefix is `lns4`.
- Go to `https://lns4-admin.sharepoint.com` to create an app catalog site.
- Left sidebar > More features > Apps > Open.

  <img width="1200" alt="image" src="screenshots/admin-center-more-features.png">
- SharePoint app catalog is used to manage and deploy SharePoint Framework solutions.
  
  <img width="1200" alt="image" src="screenshots/app-catalog-site.png">

### Create a new site collection
- You also need a site collection and a site for your testing.
- Go to `https://lns4-admin.sharepoint.com` to create a new site collection.
- In the left sidebar, select Sites > Active sites > Create.
  
  <img width="1200" alt="image" src="screenshots/create-site.png">
- Create team site.
  
  <img width="550" alt="image" src="screenshots/create-team-site.png">
- Enter required details.

  <img width="1200" alt="image" src="screenshots/site-name.png">

  <img width="600" alt="image" src="screenshots/site-owners-members.png">
- Browse to the site collection you just created: <Your site> > Site address > Click the link.
  ```http request
  https://lns4.sharepoint.com/sites/mytest/SitePages/CollabHome.aspx
  ```
- Edit site details [using this guide](https://support.microsoft.com/en-us/office/change-a-site-s-title-description-logo-and-site-information-settings-8376034d-d0c7-446e-9178-6ab51c58df42).

### SharePoint Workbench
SharePoint Workbench is a developer design surface that enables you to quickly preview and test web parts without deploying them in SharePoint. 

You can access the Hosted SharePoint Workbench from any SharePoint site in your tenancy by browsing to the following URL:
```
https://your-sharepoint-site/_layouts/workbench.aspx
# For my "mytest" site:
https://lns4.sharepoint.com/sites/mytest/_layouts/15/workbench.aspx
```
<img width="1000" alt="image" src="screenshots/site-workbench.png">

## Set up your SharePoint Framework development environment
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment

### Install Node
https://github.com/akhanalcs/nodejs-in-azure/blob/main/nodejs-learn/README.md#install-nodejs-using-nvm

Currently, I already have the latest version of Node.js installed on my machine.
```bash
$ nvm current
v23.10.0
$ node --version
v23.10.0
```

But SharePoint Framework `v1.21.*` is supported on `Node.js v22 LTS` (aka Jod), so I'll have to install that version.

<img width="800" alt="image" src="screenshots/download-nodejs.png">

```bash
$ nvm install 22
$ nvm current
v22.15.0
$ node -v
v22.15.0
```

### Install a code editor
I'm using JetBrains Rider.

TypeScript is the primary language for building SharePoint client-side web parts.

### Install development toolchain prerequisites
The SharePoint Framework development and build toolchain leverages various popular open-source tools.
While most dependencies are included in each project, you need to install a few dependencies globally on your workstation.

#### Install Gulp
Gulp is a JavaScript-based task runner used to automate repetitive tasks.
SharePoint Framework build toolchain uses Gulp tasks to build projects, create JavaScript bundles, and the resulting packages used to deploy solutions.
[FCC Video](https://youtu.be/LYbt50dhTko?si=vd9xzvy2zhC4_lDl).

```bash
npm install gulp-cli --global
$ gulp -v
CLI version: 3.0.0
Local version: Unknown
```

#### Install Yeoman
Yeoman is a scaffolding tool for modern web applications. 

Yeoman helps you kick-start new projects, and prescribes best practices and tools to help you stay productive.  
SharePoint client-side development tools include a Yeoman generator for creating new web parts.  
The generator provides common build tools, common boilerplate code, and a common playground website to host web parts for testing.

```bash
$ npm install yo --global
$ yo --version
5.1.0
```

#### Install Yeoman SharePoint generator
The Yeoman SharePoint web part generator helps you quickly create a SharePoint client-side solution project with the right toolchain and project structure.
```bash
# The @ symbol in npm package names indicates a scoped package. Scopes serve as namespaces for packages and are typically used by organizations to group related packages together.
# The @microsoft scope indicates that the package is maintained by Microsoft. generator-sharepoint is the specific package within that scope.
$ npm install @microsoft/generator-sharepoint --global

$ yo @microsoft/sharepoint --help

# Check version. More info: https://tahoeninja.blog/posts/how-to-tell-what-version-of-the-spfx-yeoman-generator-is-installed/
$ npm list -g @microsoft/generator-sharepoint
/Users/ashishkhanal/.nvm/versions/node/v22.15.0/lib
└── @microsoft/generator-sharepoint@1.21.0

$ npm ls -g
/Users/ashishkhanal/.nvm/versions/node/v22.15.0/lib
├── @microsoft/generator-sharepoint@1.21.0
├── corepack@0.32.0
├── gulp-cli@3.0.0
├── npm@10.9.2
└── yo@5.1.0
```

[More info](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/yeoman-generator-for-spfx-intro).

## Build your first SharePoint client-side web part (Hello World part 1)
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part

Client-side web parts are client-side components that run in the context of a SharePoint page. 
Client-side web parts can be deployed to SharePoint environments that support the SharePoint Framework. 
You can also use modern JavaScript web frameworks, tools, and libraries to build them.

### Create a new web part project
```bash
$ mkdir hello-world
$ cd hello-world/
$ pwd
/Users/ashishkhanal/RiderProjects/sharepoint/hello-world
```

Create a new project by running the Yeoman SharePoint Generator from within the new directory you created:
```bash
$ yo @microsoft/sharepoint

     _-----_     ╭──────────────────────────╮
    |       |    │ Welcome to the Microsoft │
    |--(o)--|    │      365 SPFx Yeoman     │
   `---------´   │     Generator@1.21.0     │
    ( _´U`_ )    ╰──────────────────────────╯
    /___A___\   /
     |  ~  |     
   __'.___.'__   
 ´   `  |° ´ Y ` 

See https://aka.ms/spfx-yeoman-info for more information on how to use this generator.
Let's create a new Microsoft 365 solution.
? What is your solution name? hello-world
? Which type of client-side component to create? WebPart
Add new Web part to solution hello-world.
? What is your Web part name? HelloWorld
? Which template would you like to use? No framework

      _=+#####!       
   ###########|       .------------------------------------.
   ###/    (##|(@)    |          Congratulations!          |
   ###  ######|   \   |  Solution hello-world is created.  |
   ###/   /###|   (@) |   Run gulp serve to play with it!  |
   #######  ##|   /   '------------------------------------'
   ###     /##|(@)    
   ###########|       
      **=+####!  
```

### Check out the code
#### README.md
Update this file if you put this in a public repo.
For testing/ learning, I won't bother updating this file.

#### package-solution.json
The `package-solution.json` contains the key metadata information about your client-side solution package and is referenced when you run the 
`package-solution` gulp task that packages your solution into an `.sppkg` file. [Reference](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/provision-sharepoint-assets#:~:text=The%20package%2Dsolution.json%20contains%20the%20key%20metadata%20information%20about%20your%20client%2Dside%20solution%20package%20and%20is%20referenced%20when%20you%20run%20the%20package%2Dsolution%20gulp%20task%20that%20packages%20your%20solution%20into%20an%20.sppkg%20file.).

`package-solution.json` determines how your solution will be packaged and how it will appear in the SharePoint App Catalog. 
It will also control how the app will get deployed, and what permission requests and isolation your app will need. [Reference](https://pnp.github.io/blog/post/spfx-21-professional-solutions-superb-solution-packages/).

1 solution can contain multiple web parts.
#### sass.json
This sass.json file is a configuration file for the SASS compilation task.  
If you visit the schema location, you can see the properties that are available to configure.
```
https://developer.microsoft.com/json-schemas/core-build/sass.schema.json
```

For eg:
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/sass.schema.json",
  "useCSSModules": true,
  "dropCssFiles": false
}
```

My current `sass.json` only contains a reference to the schema.  
It is a minimal configuration file that:
- References the schema for editor validation and IntelliSense.
- Uses default values for all SASS compilation settings.

#### gulpfile.js
The `gulpfile.js` file is the main entry point for Gulp tasks in your project.
It defines the tasks that Gulp will run when you execute commands in the terminal.
The `gulpfile.js` file is generated by the Yeoman generator and includes a set of default tasks for building, serving, and packaging your SharePoint Framework solution.

```bash
$ gulp --tasks
Tasks for ~/RiderProjects/sharepoint/hello-world/gulpfile.js
├── clean
├── build
├── default
├── bundle
├── deploy-azure-storage
├── package-solution
├── test
├── serve-deprecated
├── trust-dev-cert
├── untrust-dev-cert
├── test-only
└── serve
```

```javascript
// Node.js looks at the @microsoft/sp-build-web package's package.json file (hello-world/node_modules/@microsoft/sp-build-web/package.json)
// For CommonJS require() calls, it uses the path specified in "exports" > "." > "require" which points to "./lib-commonjs/index.js".
// A bare import like require('@microsoft/sp-build-web') matches the "." entry
// A subpath import like require('@microsoft/sp-build-web/lib/utils') would match the "./lib/*" entry
// More info here: https://nodejs.org/api/packages.html#package-entry-points

// This loads and executes that JavaScript file, which contains the actual implementation
// The build variable below receives whatever that module exports (all the classes, functions, etc.)

// The TypeScript declaration file (.d.ts) is specified in "exports" > "." > "types" which points to "./lib-dts/index.d.ts"

// How it works:
// Node.js executes the hello-world/node_modules/@microsoft/sp-build-web/lib-commonjs/index.js file
// The file creates an exports object with several properties
// The _export function in index.js is a helper that defines getter properties on an object. It's adding these properties to the module.exports object:
// initialize, reload, rig, trustDevCert, untrustDevCert
// Rather than directly assigning values, it uses getters that return the actual values when accessed.

const build = require('@microsoft/sp-build-web');

// ...

// build.rig... // "rig" comes from `const rig = new _SPWebBuildRig.SPWebBuildRig();` so check out SPWebBuildRig.js file for more info

// Passing your gulp instance (gulpfile.js) to the build system by calling the initialize function
// Setting up all SPFx build tasks that will be available to your project
// The function returns the rig, though your gulpfile doesn't capture this return value
build.initialize(require('gulp'));
```

When you run a command like `gulp bundle`:
- Gulp looks up the task named "bundle".
- It finds the task executor defined in `getTasks()` method in `SPWebBuildRig.js`.
- The executor runs `getBundleTask()`, which chains together multiple sub-tasks.
- These sub-tasks handle different aspects of the build process (TypeScript compilation, webpack bundling, etc.)

The code in `node_modules/@microsoft/sp-build-web` is primarily compiled JavaScript with TypeScript definition files (.d.ts). 
This makes it difficult to navigate to actual implementations so don't spend time looking at it. 

##### TS declaration file
The `hello-world/node_modules/@microsoft/sp-build-web/lib-dts/index.d.ts` file is a TypeScript declaration file. 
It provides type definitions for JavaScript code without containing any implementation.
These files are crucial for TypeScript's type checking system.

For eg:  
The `index.d.ts` file provides type definitions for the SharePoint Framework build system. When you import `@microsoft/sp-build-web` in a
TypeScript file, TypeScript will use these definitions to verify you're using the API correctly.

For example, in your gulpfile, the line:
```javascript
const build = require('@microsoft/sp-build-web');
// ...
```

If this were TypeScript code using imports instead, TypeScript would know that:
- `build.rig` is of type `SPWebBuildRig`. Shown in the `index.d.ts` file as `export declare const rig: SPWebBuildRig;`.
- `build.initialize()` takes a gulp instance and returns an `SPWebBuildRig`.

###### Example of how it helps
Without the `.d.ts` file:
```typescript
import * as build from '@microsoft/sp-build-web';
build.rigTypo.getTasks(); // Error only at runtime
```

With the `.d.ts` file:
```typescript
import * as build from '@microsoft/sp-build-web';
build.rigTypo.getTasks(); // TypeScript error: Property 'rigTypo' does not exist on type...
```

The declaration file enables compile-time type checking without affecting the actual JavaScript execution.

#### Other files
Check out other files.

https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part#web-part-project-structure

### Preview the web part
You can preview and test your client-side web part in the SharePoint **hosted workbench** without deploying your solution to SharePoint. 
This is done by starting a local web server the hosted workbench can load files from using the gulp task **serve**.

#### [Trusting the self-signed developer certificate](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment#trusting-the-self-signed-developer-certificate)
```bash
$ gulp trust-dev-cert
Build target: DEBUG
[21:05:54] Using gulpfile ~/RiderProjects/sharepoint/hello-world/gulpfile.js
[21:05:54] Starting 'trust-dev-cert'...
[21:05:54] Starting gulp
[21:05:54] Starting subtask 'trust-cert'...
[21:05:54] [trust-cert] Attempting to trust a development certificate. This self-signed certificate only points to localhost and will be stored in your local user profile to be used by other instances of debug-certificate-manager. If you do not consent to trust this certificate, do not enter your root password in the prompt.
Enter your password: 
[21:06:36] Finished subtask 'trust-cert' after 42 s
[21:06:36] Finished 'trust-dev-cert' after 42 s
[21:06:37] ==================[ Finished ]==================
[21:06:37] Project hello-world version:0.0.1
[21:06:37] Build tools version:3.19.0
[21:06:37] Node version:v22.15.0
[21:06:37] Total duration:45 s
```
The client-side toolchain uses HTTPS endpoints by default.

It's necessary to do this so that whenever we're referencing SharePoint online, there won't be a violation between http and https.
And the browser is able to load the development time assets from localhost.

#### [Set the SPFX_SERVE_TENANT_DOMAIN environment variable](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment#set-the-spfx_serve_tenant_domain-environment-variable-optional)
Open your shell profile and set the `SPFX_SERVE_TENANT_DOMAIN` environment variable.
```bash
$ vim ~/.bash_profile
```
Add this:
```
# Set the SharePoint Framework Hosted Workbench Test Site
export SPFX_SERVE_TENANT_DOMAIN=lns4.sharepoint.com
```

Source the profile:
```bash
$ . ~/.bash_profile
```

#### Start the local web server & launch the hosted workbench
```bash
$ gulp serve
```

- Search the web part name in the search box.

  <img width="800" alt="image" src="screenshots/hello-world-web-part.png">
- Click the web part to add it to the page.

  <img width="800" alt="image" src="screenshots/hello-world-web-part-loaded.png">
- Click the pencil icon to edit the web part. You can play with the properties.

  <img width="1000" alt="image" src="screenshots/hello-world-web-part-edit-properties.png">

### Configure the Web part property pane
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part#configure-the-web-part-property-pane

The flow looks like this (AI generated):
1. App starts → imports functions like PropertyPaneTextField, PropertyPaneCheckbox etc. from `@microsoft/sp-property-pane`
2. `IHelloWorldWebPartProps` interface defines the property structure and types
3. Default values are loaded from manifest.json's `properties` section
4. Web part class instance is created with those default values
5. `render()` method uses `this.properties.X` to display values
6. User clicks "Edit" on the web part → property pane appears
7. Property pane shows UI controls defined in `getPropertyPaneConfiguration()`
8. When user changes values in these controls, the corresponding `this.properties.X` values update
9. Web part re-renders to reflect changes

For more info, look at the code and the comments I've put there.

<img width="1200" alt="image" src="screenshots/new-properties-to-pane.png">


