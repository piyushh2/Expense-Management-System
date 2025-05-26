# Expense Management System
## Overview
The Expense Management System is a modern web application built using the SharePoint Framework (SPFx), React with TypeScript, and Material UI.  
Designed to integrate seamlessly with SharePoint Online, this application enables users to track and manage their expenses and income efficiently within a SharePoint environment.  
It offers a user-friendly interface for logging financial transactions, categorizing expenses, setting budgets, and generating reports, all while leveraging the power of SharePoint for data storage and collaboration.  

## Features

Expense and Income Tracking: Record financial transactions with details such as date, category, description, and amount.  

Category Management: Organize transactions into customizable categories (e.g., Food, Transport, Bills).  

Budget Planning: Set budgets for different categories to monitor and control spending.  

Reports and Visualizations: Generate summaries and visualize spending patterns using Material UI components (e.g., tables, charts).  

SharePoint Integration: Store data in SharePoint lists and leverage SharePoint's security and collaboration features.  

User Authentication: Utilize SharePoint's authentication for secure access, supporting single sign-on (SSO).  

Responsive Design: Built with Material UI for a consistent, responsive experience across devices.  


## Technologies Used

Frontend: React (with TypeScript), Material UI  

Framework: SharePoint Framework (SPFx)  

Data Storage: SharePoint Lists and Library  

Development Tools: Node.js, Yeoman, Gulp, TypeScript

## Prerequisites  

Before setting up the project, ensure you have the following:  



Node.js (v16.x or compatible with your SPFx version): Download  


Yeoman and Gulp CLI:  
npm install -g yo gulp-cli  



SharePoint Framework Yeoman Generator:  
npm install -g @microsoft/generator-sharepoint  



A SharePoint Online tenant with permissions to deploy SPFx solutions.  

A code editor like VS Code.  

Access to a SharePoint site for deploying the web part.  


## Installation
Follow these steps to set up the project locally and deploy it to SharePoint:  


Clone the Repository:  

git clone https://github.com/piyushh2/Expense-Management-System.git  

cd Expense-Management-System  



Install Dependencies:  

npm install  



## Configure SharePoint Connection:

Update the SharePoint configuration in config/serve.json with your SharePoint site URL:  

{
  "initialPage": "https://your-tenant.sharepoint.com/sites/your-site/_layouts/workbench.aspx"
}  



Ensure you have access to the SharePoint site and the necessary permissions.  



## Set Up SharePoint Lists:  


Create SharePoint lists (Expenses, Requests, Approval History) with columns for:
Amount (Number)  

Category (Choice or Text)  

Date (DateTime)  

Type (Choice: Expense/Income)  



Update the list name and column mappings in the SPFx web part code (e.g., in src/webparts/expenseManagement/services/spService.ts).  



## Run the Application Locally:  


Start the local development server : gulp serve  



Open your browser and navigate to the SharePoint workbench  

(e.g., https://your-tenant.sharepoint.com/sites/your-site/_layouts/workbench.aspx).  

Add the Expense Management web part to the workbench to test the application.  



## Build and Deploy to SharePoint:  


Bundle and package the solution:gulp bundle --ship  

gulp package-solution --ship  



Locate the generated .sppkg file in the sharepoint/solution folder.  

Upload the .sppkg file to your SharePoint tenant's App Catalog.  

Deploy the app to your SharePoint site and add the web part to a page.  




## Usage  

### Access the Web Part  


Navigate to your SharePoint site and edit a page.  

Add the Expense Management web part to the page.  


### Manage Transactions  


Use the web part to add, edit, or delete expenses and income.  

Select categories, input amounts, and specify dates through the Material UI-based interface.  


### View Reports  


Access the dashboard within the web part to view spending summaries and visualizations (e.g., charts or tables).

### Admin Features

Admins with appropriate SharePoint permissions can manage categories and view all user transactions via SharePoint list permissions.

## Project Structure  


src/webparts/expenseManagement/: Contains the SPFx web part code.  

src/webparts/expenseManagement/components/: React components built with TypeScript and Material UI.  

config/: SPFx configuration files (e.g., package-solution.json, serve.json).  


## Contributing
We welcome contributions to enhance the Expense Management System! To contribute:

Fork the repository.
Create a new branch:git checkout -b feature/your-feature


Make your changes and commit:git commit -m "Add your feature"


Push to the branch:git push origin feature/your-feature


Open a pull request with a detailed description of your changes.

Please ensure your code adheres to TypeScript linting rules and Material UI best practices. Include relevant tests if applicable.
## License
This project is licensed under the MIT License.
## Contact
For questions or suggestions, feel free to reach out:

#### GitHub: piyushh2
#### Email: [satijapiyush11@gmail.com]

Happy expense tracking with SharePoint!
