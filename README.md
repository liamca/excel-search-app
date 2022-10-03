# How to Create an Excel Search Application

Search applications are typically supported through Web applications, however, these can often be time consuming to create and often the people that need to create these applications may not have the development background or time to create these applications. Furthermore, often there may be times where a browser is not the preferred place to work with the data and it may be productivity tools such as Excel that are more preferrable. This tutorial is intended to show how to create a search applicaiton from within Excel that requires no-coding (other than some Power Query scripting) with very little effort. Since this tutorial leverages Power Query, with very few changes it could be adapted to provide the same functionality in Power BI.

## What is Needed
- This tutorial leverages [Azure Cognitive Search](https://learn.microsoft.com/azure/search/search-what-is-azure-search) as the search engine and you will require an Azure Subscription ([Free Subscription](https://azure.microsoft.com/free/))
- Excel 

## Getting Started
### Create the Azure Cognitive Search Index
For this tutorial we will use a sample index that comes with Azure Cognitive Search called Hotels, which is a data source consisting of fictitious hotel data. We will use this dataset to allow us to search and filter the hotel rooms within this search index.

Follow the instructions in [this page](https://learn.microsoft.com/en-us/azure/search/search-get-started-portal) to create and explore the hotel index from the Azure Portal.

### Enable CORS
Since the Excel application will connect directly to the search index, we need to enable CORS for this index. To do this, in the Azure Portal choose the "Indexes" tab and click on the "hotels-sample-index". Click "CORS" and select "All" as the allowed origin type and choose "Save".

[INSERT CORS IMAGE]

### Create a Query API Key
All queries to Azure Cognitive Search need to be authenticated. Since we will be querying from Excel, we will want to create a Query API key as this key has limited priviledges to the search index and also sufficient priviledges to suppor the Excel search application. To create this key, choose the "Keys" option from the main page of your Azure Cognitive Search service.

Under "Manage Query Keys" choose "Add". Name the key "excel-search-app" and choose "Done".

Copy the resulting Query Key and save this for a future step.

[INSERT Query Key IMAGE]

### Create the Excel Search Application
Now that we have a search index, we will create a new Excel spreadsheet 
