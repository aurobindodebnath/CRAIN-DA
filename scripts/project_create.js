let $ = require('jquery')
let fs = require('fs')
let path = require('path')


//creating project
document.getElementById('create').addEventListener('click', () => {
   let client = $('#client').val()
   let category = $('#category').val()
   let app = $('#app').val()
   let year = $('#year').val()
   let month = $('#month').val()
   let number = 1
   var filename = client+'_'+app+'_'+category+'_'+month+'_'+year+'_'+number+'.json'
   let json_obj = {
	    "client": client,
      "app":app,
	    "category": category,
	    "month": month,
	    "year": year,
      "img_limit":"2",
      "stock":[],
      "obs": {}
	   }
   let json_string = JSON.stringify(json_obj)
   while(fs.existsSync(path.join('.','resources','app','projects',filename)))
   {
     number++;
     filename= client+'_'+app+'_'+category+'_'+month+'_'+year+'_'+number+'.json'
   }
   if(!fs.existsSync(path.join('.','resources','app','projects',filename))) {
      fs.writeFile(path.join('.','resources','app','projects',filename), json_string, (err) => {
         if(err)
            console.log(err)
      })
   }
  if(filename != undefined && filename != null) {
    document.getElementById('move').href = path.join('.','templates','checklist.html')+"?q="+category+"&&f="+filename
    document.getElementById('move').click()
  }
})
