Good old CSP is well equipped to produce HTML tables accepted from EXCEL as input.  
With modern Browsers you don't even need <head> and  <body> tags.  
So the required code around your SQL result set is really slim.  
And you are free to add any formatting you need either by HTML or in SQL.  

The final trick to move your table from browser to EXCEL:  
In the method, OnPreHTTP inherited from %CSP.Page you  
set %response.ContentType="application/vnd.ms-excel"  

Now when you call the class with your browser you get asked to open or to save it.   
Next, because the extension is .cls you get asked for the program to open it.  
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

## Test and Demo
- the screenshots were created when I still used MS EXCEL and MS InternetExplorer on WIN 10
- the initial design expected a namespace **SAMPLES** (as in Class Docs)
- for testing in Docker a dedicated demo class is now created

## Docker    
Container build and start runs ALL installation steps.    
It is immedeatly ready for use as described    

### Prerequisites
Make sure you have [git](https://git-scm.com/book/en/v2/Getting-Started-Installing-Git) and [Docker desktop](https://www.docker.com/products/docker-desktop) installed.
### Installation
Clone/git pull the repo into any local directory
```
$ git clone https://github.com/rcemper/Light-weight-EXCEL-download.git
```
```
$ docker compose up -d && docker compose logs -f
````
### Test
from the terminal create the demo file by
```
USER>write ##class(DC.EXCELlight).toFile()
```
see the result by   
```
USER>$cat DC.EXCEL.html
```

Testing in browser depends on the ability of your browser to   
match with ContentType="application/vnd.ms-excel" otherwise   
you just get an offer to download the actual page:   
This may sound obvious. Though ....   
Feedback to the package shows the need for this clarification.   

This is the link:   
[http://localhost:42773/csp/user/DC.EXCELlight.cls](http://localhost:42773/csp/user/DC.EXCELlight.cls) 
