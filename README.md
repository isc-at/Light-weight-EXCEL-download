<div align="right"><h3>*** archived ***</h3></div>The full story is available on [DeveloperCommunity](https://community.intersystems.com/post/light-weight-excel-download)
But here's the light weight export to EXCEL.

Good old CSP is well equipped to produce HTML tables accepted from EXCEL as input.  
With modern Browsers you don't even need <head> and  <body> tags.  
So the required code around your SQL result set is really slim.  
And you are free to add any formatting you need either by HTML or in SQL.  

The final trick to move your table from browser to EXCEL:  
In the method OnPreHTTP inherited from %CSP.Page you  
set %response.ContentType="application/vnd.ms-excel"  

Now when you call the class with your browser you get asked to open or to save it.   
Next , because the extenison is .cls you get asked for the program to open it.  
like this: 
- ![](https://openexchange.intersystems.com/mp/img/packages/827/screenshots/4u4gejtzxfceqa55qerjca1kjzk.jpg)

And if you select EXCEL (or any compatible tool) the table is ready for the user to work with it.  
as this: 
- ![](https://openexchange.intersystems.com/mp/img/packages/827/screenshots/qduzy2m93pbx4b7revdiukhqxwe.jpg)  

# Summary:

This could be a slim solution for rater static SQL queries.   
Well suited to serve a wide distributed population of users.  

- No need for additional EXCEL installation on Caché server  
- No need for new Caché version + license upgrades to run ZENreports  
- No need for extra transport to move results to users   
- No need for local installed installed software (Squirrel)  
- No need for additional management of SQL access rights  

Rather small size of code with simple structure  

Now you will understand why I titled it "Light Weight"  

[Article in DC](https://community.intersystems.com/post/light-weight-excel-download)
