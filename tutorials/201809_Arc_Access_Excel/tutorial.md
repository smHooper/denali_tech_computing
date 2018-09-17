Access/Excel to/from ArcGIS
==============
</br>

> 
>These notes and tutorial walk through some of the methods available for viewing Access/Excel data in ArcMap and some approaches for utilizing those data (i.e., symbolization and analysis). Some familiarity with ArcGIS and Access is recommended.
>

__Contact__: samuel_hooper@nps.gov

</br>

If you'd like to use this document as a tutorial, first either [download the sample data](#download_instructions) (easy) or [clone](https://git-scm.com/docs/git-clone) this repository (advanced). Once you've done that, open __201809_arc_access_excel_tutorial.mxd__ and follow any instructions given in blocks like this:

**_Example tutorial instruction block_**
```no-highlight
1. Read this tutorial
2. Follow instructions in blocks like this
```

</br>

## Table of Contents
1. [Downloading the sample data](#download_instructions)
2. [File types](#file_types)
3. [Limitations](#limitations)
    - [Field naming restrictions](#naming_conventions)
    - [Editing a database in a non-native application](#editing_db)
4. [Reading external data in ArcMap](#reading_data)
    - [Text files](#read_text_files)
    - [Connecting to an Access database](#connect_db)
5. [Displaying/analyzing external data in ArcMap](#displaying_analyzing)
    - [Using point data](#point_data)
    - [SQL queries](#sql_queries)
    - [Joining tables](#joins)
    - [Saving the results of joins](#saving_join_results)
    - [Using spatial queries](#spatial_queries)
6. [SQLite + QGIS](#sqlite)

<a name="download_instructions"/>

</br>
</br>

Downloading the sample data
----------------------------
1. Click on [this link](https://github.com/smHooper/denali_tech_computing/blob/master/tutorials/201809_Arc_Access_Excel/201809_arc_access_excel_sample_data.zip) to go to the sample data page
3. Click the __Download__ button
4. Unzip the downloaded file and save it to a location where there are no spaces or special characters in the preceding path (e.g., C:\users\my_name\computing_tutorials)

<a name="file_types"/>

</br>
</br>

File types that are readable across platforms
-----------------------------------------------
#### .mdb
This was originally a database file for versions of Access before 2007. It's also the file extension ESRI uses for [personal geodatabases](http://desktop.arcgis.com/en/arcmap/latest/manage-data/administer-file-gdbs/personal-geodatabases.htm), so .mdb files are theoretically directly reabable in ArcGIS applications.

#### .accdb
Access 2007 introduced this file type, and it has since been the default Access file format. It is not directly reabable in ArcGIS applications, but you can establish an OLE (Object Linking and Embedding) connection to read an .accdb in ArcMap.

#### .dbf
The .dbf is a legacy file format used for storing tabular data. The table of a shapefile is stored in a .dbf and it can be read by both Access (via import) and Excel. 

#### .xls and .xlsx
Excel files are directly readable in all versions of ArcGIS 10.x, although this functionality was a bit buggy in earlier versions.

#### other formats
Delimited text files (.csv, .tsv, or .txt) are the most universal file format. You can also connect to other databases that conform the the ODBC (Object Database Connectivity) standard (e.g., [Postgres](https://www.postgresql.org/), [SQLite](https://www.sqlite.org/index.html))


<a name="limitations"/>

</br>
</br> 

Limitations
------------

<a name="naming_conventions"/>


### Field naming restrictions
ArcMap and Access enforce different restrictions on field (column) names. In ArcMap, fields must:
- start with a letter
- contain only letters, numbers, and underscores
- have <= 64 characters

Although Access doesn't place the same restrictions on field names, these are generally safe rules to follow when creating tabular data in any software package. The only restrictions on fields in data created in ArcMap to be read in Access is they can't be named with any of the [Access reserved words](https://support.microsoft.com/en-us/help/286335/list-of-reserved-words-in-access-2002-and-in-later-versions-of-access).

<a name="editing_db"/>

### Editing a database in a non-native application
Theoretically, these methods for utilizing data across software packages should only be for viewing data, not editing or modifying a database. Because a table from an Access database doesn’t have an ObjectID field created and managed by ArcGIS, the data can’t be selected on the map (although records in the table can be) nor can they be edited without exporting and making a copy of the data. 

Access also can do some unpredictable things to database files like adding system and extra tables without warning. It also can't maintain any of the spatial information in a database created in ArcGIS applications, which can get lost when editing a database in Access. For these and other reasons, ESRI strongly discourages modifying a personal geodatabase in Access.

<a name="reading_data"/>

</br>
</br>

Reading external data in ArcMap
---------------------------------

<a name="read_text_files"/>

### Text files and Excel worksheets
These can be read directly in ArcMap by any typical method for adding data (e.g., the Add Data button, dragging and dropping from ArcCatalog). *NOTE: When you add tabular data, the Table of Contents automatically changes to the “List by Source” view, and if you switch back to the standard “List by Drawing Order” view, the tabular data are no longer visible.*

<a name="connect_db"/>

### Connecting to an Access database
Since .accdb files aren't directly readable in ArcMap, you need to establish a connection to the database:

1. Add the __Add OLE DB connnection__ tool to ArcCatalog
    - Open __ArcCatalog__
    - From the main menu, select __Customize__ > __Customization Mode...__
    - From the Customize window, select the _Commands_ tab
    - Type __"connection"__ in the *Show commands containing* text field. The only option in the __Commands__ list should be *Add OLE DB Connection*
    - ![alt text](https://github.com/smHooper/denali_tech_computing/blob/master/tutorials/201809_Arc_Access_Excel/images/add_ole_db_button.PNG "Add the 'Add OLE DB Connection' tool")
    - Drag the __Add OLE DB Connection__ tool to any toolbar (underneath the main menu)
2. Click the __Add OLE DB Connection__ tool you just added to a toolbar. In the dialog that appears:
    - Under the __Provider__ tab, select __Microsoft Office 12.0 Access Database Engine OLE DB Provider__ and cleck __Next >>__
    - In the __Connection__ tab under item 1, enter the full path to the .accdb file in the *Data Source* text field. *NOTE: A quick method for copying the full path of a file on Windows is to open a Windows Explorer window and hold the Shift key and right-click. Then select "Copy as path". This will copy the path with quotation marks around it, so you'll have to delete those once you've pasted it into the Data Source text field.*
    - ![alt text](https://github.com/smHooper/denali_tech_computing/blob/master/tutorials/201809_Arc_Access_Excel/images/data_link_properties__select_db.png "Enter the Data Source")
    - Click __Test Connection__
    - Click __OK__ if the connection was successful
3. Load a table from the connection
    - In __ArcCatalog__, expand __Database Connections__. If you don't see an item called something like “OLE DB Connection”, right-click __Database Connections__ and click __Refresh__
    - Rename the connection you just made to something identifiable
    - Double-click the connection to expand it, and drag a table form it into ArcMap  (you can’t add the entire connection, just one table at a time)
    

Although .mdb files can be read directly by ArcGIS, ESRI recommends reading any tables in ArcMap using an OLE connection. You can connect to a .mdb with the same steps as above for a .accdb, in Step 2 select __Microsoft Jet 4.0 OLE DB Provider__ in the __Provider__ tab.

**_Tutorial: Connect to an Access database_**
```no-highlight
1. Follow the instructions above under Step 1 for adding the 'Add OLE DB connection' button
2. Click the 'Add OLE DB Connection button'
3. In the Provider tab of the window that appears, select 'Microsoft Office 12.0 Access Database 
Engine OLE DB Provider' and click 'Next >>'
4. Open a Windows Explorer window and navigate to the location where you saved the sample data 
from this repository 
5. Hold down the shift key and right-click on PatrolOverflightObsDb.accdb
6. Select "Copy as path"
7. In ArcCatalog, select the 'Connection' tab
8. Paste the path into the 'Data Source' text field
9. Delete the quotations marks at the beginning and end of the path you just pasted.
10. Test the connection and click OK if it was successful
11. In ArcCatalog, right-click on 'Database Connections' and select 'Refresh'
12. Right-click on the connection titled "OLE DB Connection", select 'Rename', and type 
'overflights'
13. Double-click the 'overflights' connection to expand it. 
14. Drag and drop the table 'tbl_EventsPerUnit' in the ArcMap document.
```

<a name="displaying_analyzing"/>

</br>
</br>

Displaying and analyzing data from Access/Excel in ArcMap
------------------------------------------------------------

<a name="point_data"/>

### Point data with lat/lon
Point data from any table can be displayed in ArcMap:
1. Right-click on the table in the __Table of Contents__
2. Select __Display XY Data__
3. Select the fields containing lat and lon values
4. If the coordinates from the table are in a different coordinate system than the current projection of the Data Frame, select __Edit__ in the lower right corner and select the coordinate system of the table's coordinates

**_Tutorial: Display point data_**
```no-highlight
1. In ArcCatalog, find the Excel workbook called points.xlsx. Double-click it to expand it.
2. Drag and drop the 'event_locations' table into the ArcMap document
3. Right-click on the table in the 'Table of Contents' and select 'Display XY Data'
4. Click the 'Edit...' button in the bottom right corner.
5. Select 'WGS 1984' as the coordinate system (Geographic Coordinate Systems > World> WGS 1984) 
and click 'OK'
6. Click 'OK' and click 'OK' again in the alert that appears
```

<a name="sql_queries"/>

</br>

### SQL queries in ArcMap
You can perform SQL-like SELECT queries on any table in ArcMap. To do so:
1. Open the table
2. Click on the __Table Options__ button in the top left corner of the table's menu
3. Select __Select By Attributes...__
*Note that you are only writing the WHERE clause of the SELECT statement*

**_Tutorial: Select specific records using SQL_**
```no-highlight
1. Open the attribute table of the bc_units layer
2. Click on the 'Table Options' button
3. Select 'Select By Attributes...'
4. In the lower text field, type ' "Name" LIKE '%Glacier' ' to select just backcountry units that 
end with the word "Glacier"
5. Click 'Apply' and 'Close'
```

<a name="joins"/>

</br>

### Joining tables
You can use joins to append aspatial data from an Access or Excel table to spatial data like a shapefile. These tables must share a common ID field just like any relational database. To join a table to a spatial layer in ArcMap:
1. Right-click on the layer in the __Table of Contents__
2. Select __Joins and Relates__ > __Join...__
3. In the __Join Data__ window:
    - In item 2, select table you want to join to the spatial layer you right-clicked in Step 1
    - In item 1, select the field from the spatial layer containing a unique ID shared by the table you selected for item 2
    - In item 3, select the field from table containing a unique ID shared by the layer
    - Select the type of join you want under __Join Options__. For an [OUTER](https://www.w3schools.com/sql/sql_join_full.asp) join, choose __Keep all records__. For an [INNER](https://www.w3schools.com/sql/sql_join_inner.asp) join choose __Keep only matching records__.
4. Select __Validate Join__ to verify the tables can be successfully joined. Reasons a join could fail are that there are no matching records or one or more fields violate these [naming restrictions](#naming_conventions)
5. Select __OK__

Once you've joined a table to a layer, you can then use the fields from the joined table to symbolize or label the the layer. 

**_Tutorial: Join a table from an Access database_**
```no-highlight
1. Right-click on the bc_units layer and select 'Joins and Relates' > 'Join...'
2. In the item 2 dropdown menu, select the 'tbl_EventsPerUnit' table from the 'overflights' 
database connection you created eariler
3. In the item 1 dropdown menu, select 'Unit' as the field to join on from the layer
4. If 'BCUnit' is not automatically selected in item 3, select it as the field from the database 
table to join on
5. Click 'OK'. When you open the 'bc_units' attribute table, you should see the two fields from 
'tbl_EventsPerUnit'. Note that since we chose to keep all records, some records are null because 
they don't exist in the database table.
```

<a name="save_join_results"/>

</br>

### Saving results of a temporary join
Note that a join created in ArcMap is a virtual join only available within this ArcMap document (.mxd file). If you read the same shapefile the layer points to in another ArcMap document, it will not have the same table joined to it. Also note that if the connection to its parent database dies or the path to the table changes, the join will be lost. 

To save the results of a join, you can either:
1. Use the __Field Calculator__ to copy values from the joined table (preferred for saving a single field)
    - Open the attribute table of the layer
    - Click on the __Table Options__ button in the top left corner of the table
    - Select __Add Field__
    - Enter a name for the field, select the field type, and change the fields parameters (if necessary)
    - Click __OK__
    - Right-click on the field you just created
    - Select __Field Calculator...__
    - In the __Fields__ menu, find the field from the joined table you want to save values from and double-click it
    - Select __OK__
2. Export the data and save it as a new shapefile, feature class, or table (preferred for saving multiple fields)
    - Right-click on the layer in the table of contents
    - Select __Data__ > __Export data...__
    - Choose a location to save it to
    - Click __OK__


**_Tutorial: Copy the values from a joined table_**
```no-highlight
1. Click the 'Clear Selection' button (third from the right at the top of the attribute table)
2. Right-click the 'Table Options' button and select 'Add Field...'
3. Set the name of the field to 'intensity' and leave all other parameters as they are.
4. Click 'OK'
5. Right-click on the column heading of the field you just created (it will say 'bc_units.intensity' 
indicating that it belongs the bc_units table, not the joined table)
6. Select 'Field Calculator...'
7. In the 'Fields' menu, double-click 'CountOfIntensity' and click 'OK'. If you get an alert 
message, select 'Yes' indicating that you want to modify field values outside of an editing 
sesssion.
8. Right-click on the 'bc_units' layer in the 'Table of Contents' and select 'Joins and Relates' >
 'Remove join(s)' > 'tbl_EventsPerUnit'. Now if you open the attribute table again, the fields 
 from tbl_EventsPerUnit will be gone, but the 'CountOfIntensity' is permanently stored in the 
 'intensity' field of the 'bc_units' shapefile.
```

<a name="spatial_to_aspatial"/>

</br>

### Using spatial queries to inform aspatial data
The [Select Layer By Location tool](http://desktop.arcgis.com/en/arcmap/latest/tools/data-management-toolbox/select-layer-by-location.htm) is similar to SQL-like queries, but it uses spatial rather than tabular information to select particular records. You can, therefore, isolate particular records from a database table representing spatial features that share some spatial relationship with features in a spatial layer. After exporting those records as a text file, that spatial query is ready to be used in your aspatial database (via [importing external data](https://support.office.com/en-us/article/import-or-link-to-data-in-a-text-file-d6973101-9547-4315-a8f8-02911b549306)).

**_Tutorial: Select by location tool_**
```no-highlight
1. In the 'Search' window of ArcMap, type "select by" in the search text field
2. Click on the 'Select Layer By Location' tool (should be 5th down from the top)
3. Drag the 'points' layer into the 'Input Feature Layer' field
4. In the 'Relationship' field, select 'WITHIN'
5. Drag the 'bc_units' layer into the 'Selecting Features' field
6. Click the check box next to 'Invert Spatial Relationship'
7. Click 'OK'. Records that are selected are now all points that do not fall within a backcountry 
unit.
```

<a name="sqlite"/>

</br>
</br>

\*\*Open source proselytization\*\*: SQLite + QGIS makes this __WAY__ easier
----------------------------------------------------------------------------
[SQLite](https://www.sqlite.org/index.html) is sort of an open source cousin to Access in that they're both single-file databases that are relatively simple to manage. With SQLite, however, you can use the [Spatialite](https://www.gaia-gis.it/fossil/libspatialite/index) extention that enables the same spatial queries as the Select By Location tool in ArcMap. Since you can also perform the same SQL queries, it combines all of the features described above in one neat package.

Even better is the combination of SQLite/Spatialite and [QGIS](https://qgis.org/en/site/), the open source cousin to ArcMap. SQLite databases can be read directly in QGIS without a complicated process to manage database connections and without any need to ever export or make copies of your data. For a brief primer on using SQLite in QGIS see [this tutorial](https://docs.qgis.org/2.8/en/docs/training_manual/databases/spatialite.html). To try out QGIS, you can [download it for free](https://qgis.org/en/site/forusers/download.html).

If you're interested in this approach, feel free to send any questions to samuel_hooper@nps.gov.
