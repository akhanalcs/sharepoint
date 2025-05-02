# Sharepoint
Learning SharePoint.

## Sharepoint Basics
https://support.microsoft.com/en-us/office/get-started-with-sharepoint-909ec2f0-05c8-4e92-8ad3-3f8b0b6cf261

<img width="450" alt="image" src="screenshots/sharepoint-link.png">

Make sure you're logged in using your Microsoft 365 account.

### Microsoft Forms
2 ways to create forms:
- [Microsoft Forms](https://support.microsoft.com/en-us/office/create-a-form-with-microsoft-forms-4ffb64cc-7d5d-402f-b82e-b1d49418fd9d)
- [Microsoft Forms web part on a SharePoint site](https://support.microsoft.com/en-us/office/use-the-microsoft-forms-web-part-on-a-sharepoint-site-d4b4d3ce-7860-41e4-8a98-76380efe7256) (Shown below)

#### Create a form using Microsoft Forms web part on a SharePoint site
1. Go to your SharePoint site.
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
- Left sidebar > More features.
  - Locate the section Apps and select Open.

    <img width="1200" alt="image" src="screenshots/admin-center-more-features.png">
- SharePoint app catalog is used to manage and deploy SharePoint Framework solutions.
  
  <img width="1200" alt="image" src="screenshots/app-catalog-site.png">

### Create a new site collection
- Go to `https://lns4-admin.sharepoint.com` to create a new site collection.
- In the left sidebar, select Sites > Active sites.