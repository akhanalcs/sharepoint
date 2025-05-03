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

  <img width="500" alt="image" src="screenshots/site-owners-members.png">
- Browse to the site collection you just created: <Your site> > Site address > Click the link.
  ```http request
  https://lns4.sharepoint.com/sites/mytest/SitePages/CollabHome.aspx
  ```

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

<img width="700" alt="image" src="screenshots/download-nodejs.png">

```bash
$ nvm install 22
$ nvm current
v22.15.0
$ node -v
v22.15.0
```

### Install a code editor
I'm using JetBrains Rider.

### Install development toolchain prerequisites
The SharePoint Framework development and build toolchain leverages various popular open-source tools.
While most dependencies are included in each project, you need to install a few dependencies globally on your workstation.

#### Install Gulp
Gulp is a JavaScript-based task runner used to automate repetitive tasks.
SharePoint Framework build toolchain uses Gulp tasks to build projects, create JavaScript bundles, and the resulting packages used to deploy solutions.

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

### Preview the web part
You can preview and test your client-side web part in the SharePoint **hosted workbench** without deploying your solution to SharePoint. 
This is done by starting a local web server the hosted workbench can load files from using the gulp task **serve**.







