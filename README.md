# Kredo Kodo Excel Report Generation Library
A high-performance and easy to use library to create basic Excel reports by just passing in a DataTable.

## Table of Contents

* Introduction
* Benefits and Features
* Getting Started
* Roadmap
* History

### Introduction

The purpose of this project is to provide an easy to use library, in the form of a NuGet package, that provides a fast and simple way of automatically converting a DataTable into a simple Excel report.

### Benefits and Features

* Easy to use: Simple 'dot your way to success'
* High performance: thousands of rows in milliseconds
* Flexible: Create multiple worksheets in a single workbook in a customized way
* .NET Standard 2.0 Compliant

### Getting Started

Install the nuget package and new up a Report Generator as shown:

> var reportGenerator = new ReportGenerator();
OR
> Dim reportGenerator = new ReportGenerator()

Let's say you have two DataTable objects to work with representing animals held in a zoo called Monkeys and the other Kangaroos.  You might want to generate a simple Excel report that shows the number, name, sex, etc of each animal on their own worksheets to help keep track of them.

Here's an example of how to take those two DataTable objects and generate such a report:

> var reportBytes = reportGenerator
>                       .AddWorksheet(monkeyData, "Monkeys")
>                           .ConfigureColumn("name_row", "Animal Name")
>                           .ConfigureColumn("id_row", "Animal ID #")
>                           .ConfigureColumn("sex_row", "Animal Sex")
>                           .FinalizeWorksheet(12)
>                       .AddWorksheet(kangarooData, "Kangaroos")
>                           .ConfigureColumn("name_row", "Animal Name")
>                           .ConfigureColumn("id_row", "Animal ID #")
>                           .FinalizeWorksheet(11)
>                       .GenerateWorkbookBytes("ANIMALS");

Congratulations, you've just generated your very first Excel report by 'dotting your way to success'!  Now, let's break this down a little to help you understand what's going on.

__.AddWorksheet()__ - This, simply named, adds a worksheet to the workbook.  The first parameter is the DataTable that you wish you to use.  The second parameter is the name you want the worksheet to be.

__.ConfigureColumn()__ - Your DataTable might have 6 columns of data coming back from the database, but that doesn't mean that you have to use all of them.  For every column you wish to include in the worksheet just configure it and leave out the ones you don't want.  In fact, if you look at the Kangaroos section, you'll see that I left the Sex column off as I don't want it displayed for whatever reason in the report.

The first parameter is the actual name of the column stored in the DataTable.  This is usually the name of the database table's column or whatever the alias supplied was.  This string must match or an exception will be thrown.

The second parameter represents what you want the column to be named in the Excel worksheet and can be whatever you want.

Optional parameters: 
* Column Format - Simple Excel formatting like adding commas to long numbers
* Is Column Centered - Centers the data in the rows
* Column Background Color - Want all the rows in this column to be pink?  This is how you do it.

__.FinalizeWorksheet()__ - When you're done working with the column data you need to tell the Report Generator that you're done with the worksheet so you can either generate or add another.  The only optional parameter here is to set the font size of all the data within the worksheet which currently defaults to 10.

__.GenerateWorkbook() & GenerateWorkbookBytes()__ - Now is the moment you've been waiting for, actually generating the report and ensuring management get's all the juicy data they need to ensure the animals are well cared for.  

The first option returns a a ClosedXML XLWorkbook object so you can continue to manually manipulate it in any way you want.  

The second option assumes you don't want to manipulate it any further and just returns actual bytes that you can stream to browser client or do whatever else it is you do with your bits.

Finally, you can pass in an optional string parameter that will add a large header to every worksheet (above the column names).  In this case, the word ANIMALS will be displayed at the top of every worksheet.

### Roadmap

- Thinking about handling full DataSets and looping through the tables somehow.
- Perhaps adding more Custom Format options for column configurations.

If you have any ideas or feature requests please submit a pull request or comment.

### History

This library grew out of projects that I and my co-workers/friends worked on throughout 2017-2018.  We were tasked with converting manually generated XML StringBuilder files and making them comply with newer Excel XLSX standards.  In addition, formatting was really simple and so having an easier way to just throw a DataTable or two at it and get a nice looking, yet simple, report was key.
