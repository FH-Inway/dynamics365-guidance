---
title: Import the business process catalog into Azure DevOps
description: Read how you can use Microsoft's business process catalog to build an implementation project in Azure DevOps.
author: rachel-profitt
ms.author: raprofit
ms.topic: conceptual #Required; don't change.
ms.collection: #Required; Leave the value blank.
ms.date: 11/15/2023
ms.custom: bap-template #Required; don't change.
---

# Import the business process catalog into Azure DevOps

This article describes how you can use Microsoft's business process catalog as a template that you import into a project in Azure DevOps Services to manage your Dynamics 365 implementation project.

## Generate work items from the business process catalog

There are many reasons why a tool such as Azure DevOps is critical to the overall success of a Dynamics 365 implementation. By using a service such as [Azure Boards](/azure/devops/boards/get-started/what-is-azure-boards?view=azure-devops&preserve-view=true) with the business process catalog, you can accelerate the deployment and follow the recommendations in the [Process-focused solution](../implementation-guide/process-focused-solution.md) articles in the Dynamics 365 implementation guide content. The following list describes just some of the benefits of using the business process catalog to build work items to manage your Dynamics 365 implementation.

- **Efficiency and time savings**: The business process catalog provides a standardized and comprehensive list of business processes. It can save customers and partners significant time, because they don't have to research, define, and document business processes from scratch.
- **Recommended practices and industry standards**: The documentation that accompanies the catalog is available in [the Dynamics 365 guidance hub](/dynamics365/guidance/business-processes). The content of this documentation often includes recommended practices and industry-standard business processes. By applying these predefined processes, you help ensure that technology solutions are aligned with recognized industry standards and compliance requirements.
- **Reduced risk**: Use of established and standardized processes helps reduce the risk of errors and oversights in the implementation of technology solutions. These processes are tried and tested, and therefore help reduce the likelihood of costly mistakes.
- **Alignment with Microsoft technologies**: The business process catalog is designed to work seamlessly with Microsoft technologies, including Dynamics 365, Microsoft Power Platform, and Azure. This alignment can simplify integration and interoperability, so that it's easier to build and deploy technology solutions.
- **Scalability**: As businesses grow, their processes might have to evolve. The business process catalog provides a foundation that can be scaled and customized as required to ensure flexibility for future changes.
- **Community and collaboration**: Use of a standardized catalog can help foster collaboration and knowledge sharing in the Microsoft community. Customers and partners can benefit from the experiences and insights of others who use similar processes.
- **Training and onboarding**: Standardized processes can help streamline the onboarding and training of new employees or team members who join an organization. They provide a clear reference point for understanding how the organization operates.

In conclusion, the business process catalog offers customers and partners a valuable resource for efficiently and effectively implementing technology solutions in a way that's aligned with industry standards. It simplifies the process of designing, customizing, and deploying solutions, and ultimately leads to improved productivity, reduced risk, and better business outcomes.

> [!NOTE]
> This article assumes that you use [Azure Boards](/azure/devops/boards/?view=azure-devops&preserve-view=true), and that you've downloaded the business process catalog.

[!INCLUDE [daf-bus-proces-download](~/../shared-content/shared/guidance-includes/daf-bus-proces-download.md)]

## Before you import

Before you can import the project into Azure Boards, there are a few things that you must do and consider. Use the following list as a guide and checklist to ensure that you're ready to import the catalog.

1. Define your project scope.

    We recommend that you use the workbook as a starting point to define the scope. At the most basic level, delete any rows that don't apply to your project. Learn more about how to define your project scope at [Process-focused solution](/dynamics365/guidance/implementation-guide/process-focused-solution). 

1. Create a project in the Azure DevOps Services tenant.

    The template that we provide is designed to work with the *Agile* work item process type. Learn more at [Create a project in Azure DevOps](/azure/devops/organizations/projects/create-project?view=azure-devops&preserve-view=true&tabs=browser) and [Agile process work item types](/azure/devops/boards/work-items/guidance/agile-process?view=azure-devops&preserve-view=true).

1. Define area paths in the project settings.

    For each end-to-end process that is in scope, create one area path. Learn more at [Define area paths and assign to a team](/azure/devops/organizations/settings/set-area-paths?view=azure-devops&preserve-view=true&tabs=browser).

1. Add custom fields as required. The template includes four custom fields. Use the following guidance to create the fields. Alternatively, delete the columns from the template. Learn more at [Add and manage fields](/azure/devops/organizations/settings/work/customize-process-field?view=azure-devops&preserve-view=true).

    - **Business owner**: Add this field as an **Identity** field, so that you can select a user or person in the identity picker. Learn more at [Add an Identity field](/azure/devops/organizations/settings/work/customize-process-field?view=azure-devops&preserve-view=true#add-an-identity-field).
    - **Business process lead**: Add this field as an **Identity** field, so that you can select a user or person in the identity picker. Learn more at [Add an Identity field](/azure/devops/organizations/settings/work/customize-process-field?view=azure-devops&preserve-view=true#add-an-identity-field).
    - **Business outcome category**: Add this field as a **Picklist** field, so that users can select an option in a dropdown list. Learn more at [Add a picklist field](/azure/devops/organizations/settings/work/customize-process-field?view=azure-devops&preserve-view=true#add-a-picklist). We recommend that you create the following three options for the list:

        - **Business unit**: Use this option when the work item is for a specific business unit.
        - **Organization**: Use this option when the work item is for the entire organization.
        - **Process team**: Use this option when the work item is for a subset of your business unit, organization, or group of people in your organization. Although we use the term *process team*, you can use any other term that is appropriate for your project.

    - **Process sequence ID**: Add this field as a custom **Text (single line)** field, so that users can enter an ID for the process. Learn more at [Add a custom field](/azure/devops/organizations/settings/work/customize-process-field?view=azure-devops&preserve-view=true#add-a-custom-field).

1. Insert any other rows that your project requires.

    You might need more epics, features, or user stories. Epics use the first **Title** column, features use the second **Title** column, and user stories use the third **Title** column. To establish a firm relationship between the rows, don't insert the next *Epic* or *Feature* row until you've listed all rows that require a relationship to the last epic or feature. You might want to consider adding other work item types too, such as *Configuration* or *Workshops*. However, the template that we provide doesn't include other work item types.

1. Complete the other columns in the workbook as required. Use the following recommendations as guidance.

    - **Description**: Optionally add a detailed description for your business processes before you import, or work on this description throughout the project. In future releases, we plan to prepopulate this column for you. 
    - **Assigned to**: Typically, select the consultant or person who is responsible for configuring the process from the partner organization. Make sure that the person is already added to your project as a user.
    - **Business owner**: Typically, select the stakeholder from the customer organization that is responsible for the business process. Make sure that the person is already added to your project as a user.
    - **Business process lead**: Typically, select the subject matter expert from the customer organization that is responsible for the business process. Make sure that the person is already added to your project as a user.
    - **Tags**: Optionally create tags for sorting, filtering, and organizing your work items. The default template doesn't include any tags. Consider using this column to separate departments, phases, geographic regions, or product families (for example, customer engagement apps and finance and operations apps).
    - **Priority**: By default, all rows in the workbook have a priority of *1*. However, you can change the priorities to suit your needs. A priority of *1* indicates "must have" features, and a priority of *3* indicates "nice to have" features. You can also make your own definitions. In this case, we recommend that you document them for your project team.
    - **Risk**: Optionally add a rating for the risk. For example, you might give a high risk score to processes that are very complex or require lots of modification.
    - **Effort**: Optionally add a rating for the effort. For example, you might give a high effort score to processes that require integration or modification.

1. Update the **Area path** value in the file.

    You must replace the value in the **Area path** column with the exact name of your project and area paths. If you create the areas paths so that they match the end-to-end process names, you just have to replace the text *DevOps Product Catalog Working Instance* with the name of your project in your area path.

1. Optional: Add more columns to the file, or remove columns that you don't plan to use before you import. If any of the custom fields that you add to your Azure DevOps project are mandatory, make sure that you include them in the file. Otherwise, import of the file might fail.
1. Split large files for import.

    Determine whether you must split your file into multiple files for upload. Azure DevOps limits the number of rows that can be uploaded in one import to 1,000. If your final file has more than 1,000 rows, split the file. When you split the file, it's critical that all epics, features, and user stories that are related to the same end-to-end process are in the same file. For example, if row 1000 is in the middle of the [order to cash](order-to-cash-overview.md) process after the deletion and insertion of any required rows, split the file at the first row for *order to cash*. In this way, you ensure that all *order to cash* processes are included, and that you can establish the relationships during the import. If you try to import the entire catalog, you must split the file into four parts for import.

1. The file must be saved as a .csv file before it can be imported into Azure DevOps.

    If you added columns and features such as formatting or formulas in the workbook, and you don't want to lose them, consider saving a version of the file as an .xlsx file. This version can help you avoid losing those features. However, the version that you import must be the .csv file.

> [!NOTE]
> If the version of the catalog that you're about to save contains special characters such as commas (,) or quotation makes ("), remove them before you save the .csv file. For example, the October 2023 version of the uncustomized catalog contains the entry *Implement "secret shopper" program*. Change this entry to *Implement secret shopper program* before you save the .csv file.

## Import the file

After you prepare your file for import and configure the basic setup in the project with area paths, security, teams, and users, you can import your work items. Learn more at [Import update bulk work items with CSV files](/azure/devops/boards/queries/import-work-items-from-csv?view=azure-devops&preserve-view=true).

## After you import

After you import the file, validate that the import was successful. If file import fails, use the messages that are provided to fix the issue, and then try again. After the file is successfully imported, you can start to use the features of Azure Boards to manage your project. The following list includes a few tasks and tips to consider.

- [Use backlogs to manage projects](/azure/devops/boards/backlogs/backlogs-overview?view=azure-devops&preserve-view=true)
- [Implement Scrum work practices in Azure Boards](/azure/devops/boards/sprints/scrum-overview?view=azure-devops&preserve-view=true)
- [Use managed queries to list work items](/azure/devops/boards/queries/about-managed-queries?view=azure-devops&preserve-view=true)
- [Analytics & Reporting](/azure/devops/report/?view=azure-devops&preserve-view=true)
