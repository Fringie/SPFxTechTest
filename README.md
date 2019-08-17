## Scenario
The finance department of a of an organisation has a requirement to replace their current paper based expenses process with a new online version. The organisation has been using Office 365 for several years and has recently upgraded their intranet (https://modernintranet.sharepoint.com) to use Modern features. All their users are configured within Azure Active Directory and have basic profile information.

## Task 1
In preparation for the move from the paper based process to an online process, the Finance Department have defined the required data and created a backend list called “Finance Expense Claims” within the Finance Team Site to handle the storage of expense claims. The fields created for the list are:
* Firstname (text field)
* Lastname (text field)
* Expense Date (date field)
* Expense Description (multi line text field)
* Expense Cost (currency)

All the fields are required.

### Task Description
Create a SharePoint Framework web part that allows users of the Finance Department to submit their expense claims online.

### Notes
Don’t worry about connecting/submitting to an actual SharePoint site but do mock it out and show how you would.

## Task 2
Other departments wish to have their own version of the online Expense Claims form. The Finance Department have decided to create further backend lists within the Finance Team Site to store each of the other departments submissions
* HR – HR Expense Claims
* Operations – Operations Expense Claims.

### Task Description
Update your SharePoint Framework webpart to allow it to be deployed by other departments while having any submissions be stored in the correct backend list.

## Task 3
Multiple departments are now using the Expenses Claims form and performance is becoming a problem.

### Task Description
Update your SharePoint Framework Web Part to use an Office 365 CDN to deliver the assets required for the Web Part.