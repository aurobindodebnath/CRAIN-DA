/* ----------------------- REPORT GENERATING FUNCTION ------------------------------------*/

//getting the column name by number
function convert(number){
    number = number - 1
    var alphabets = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    var result = ""
    do{
        var r = Math.floor(number%26);
        var q = Math.floor(number/26);
        result = alphabets[r] + result
        number = q-1
    } while(q != 0)
    return result;
}

function createSheet(){
  console.log("creating sheet for", file_obj)
  var Excel = require('exceljs');
  var workbook = new Excel.Workbook();
  workbook.xlsx.readFile(path.join('.','resources','app','templates','base.xlsx'))
    .then(function() {
        let startrow = 10
        let rowitr=3
        let colitr=dimensions["ini_col"]
    //setting the author
        workbook.creator = "KPMG"
		
		//update the cover page
		var worksheet_cover = workbook.getWorksheet('Cover');
        let cell1 = worksheet_cover.getCell('C8').value
        cell1 = cell1.replace("var11111",file_obj["client"])
        cell1 = cell1.replace("var22222",cat[file_obj["category"]])
        worksheet_cover.getCell('C8').value = cell1
		cell1 = worksheet_cover.getCell('C12').value
        cell1 = cell1.replace("var33333",file_obj["month"])
        cell1 = cell1.replace("var44444",file_obj["year"])
        worksheet_cover.getCell('C12').value = cell1
		cell1 = worksheet_cover.getCell('C15').value
	    cell1 = cell1.replace("var44444",file_obj["year"])
        worksheet_cover.getCell('C15').value = cell1	
		
		//update disclaimer
		var worksheet_disclaimer = workbook.getWorksheet('Disclaimer and Assumptions');
		cell1 = worksheet_disclaimer.getCell('B3').value.split('var11111').join(file_obj["client"])
        worksheet_disclaimer.getCell('B3').value = cell1	
		cell1 = worksheet_disclaimer.getCell('B20').value.split('var11111').join(file_obj["client"])
        worksheet_disclaimer.getCell('B20').value = cell1	
		cell1 = worksheet_disclaimer.getCell('B22').value.split('var11111').join(file_obj["client"])
        worksheet_disclaimer.getCell('B22').value = cell1	
		cell1 = worksheet_disclaimer.getCell('B23').value.split('var11111').join(file_obj["client"])
        worksheet_disclaimer.getCell('B23').value = cell1	
		cell1 = worksheet_disclaimer.getCell('B24').value.split('var11111').join(file_obj["client"])
        worksheet_disclaimer.getCell('B24').value = cell1	
		
		//update observations & Annexures
        var worksheet = workbook.getWorksheet('Observations');
        var worksheet2 = workbook.getWorksheet('Annexures');
		cell1 = worksheet.getCell('A1').value
	    cell1 = cell1.replace("var11111",file_obj["app"])
        cell1 = cell1.replace("var22222",cat[file_obj["category"]])
        let cell2 = worksheet.getCell('J9').value
        cell2 = cell2.replace("var11111", file_obj["client"])
        worksheet.getCell('A1').value = cell1
        worksheet.getCell('J9').value = cell2
		
		//update Annexures
		cell1 = worksheet2.getCell('A1').value
	    cell1 = cell1.replace("var11111",file_obj["app"])
        cell1 = cell1.replace("var22222",cat[file_obj["category"]])
        worksheet2.getCell('A1').value = cell1	

        for(var i=0; i<file_obj["stock"].length;i++)
        {
          let ann = ""
          if(file_obj["obs"][file_obj["stock"][i]]["annexure"])
          {
            ann = "Annexure"
          }else{
            ann = "N/A"
          }
          var value = [i+1,file_obj["obs"][file_obj["stock"][i]]["observation"],file_obj["obs"][file_obj["stock"][i]]["detOb"],file_obj["obs"][file_obj["stock"][i]]["affected"],file_obj["obs"][file_obj["stock"][i]]["criticality"].toUpperCase(),file_obj["obs"][file_obj["stock"][i]]["risk"],file_obj["obs"][file_obj["stock"][i]]["recommendation"],ann]
          worksheet.spliceRows(startrow, 0, value);
          let cur_row = worksheet.getRow(startrow)
          cur_row.height = Math.min(cur_row.height, 220)
          worksheet.getCell('A'+startrow).alignment = { vertical: 'top', horizontal: 'center', wrapText: true };
          worksheet.getCell('A'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('B'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          worksheet.getCell('B'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('C'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          worksheet.getCell('C'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('D'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          worksheet.getCell('D'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('E'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('E'+startrow).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          worksheet.getCell('F'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          worksheet.getCell('F'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('G'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          worksheet.getCell('G'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('H'+startrow).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
		  worksheet.getCell('H'+startrow).font = {name: 'Univers for KPMG',size: 10, underline: true, color: { argb: 'FF0000EE' }};
		  worksheet.getCell('I'+startrow).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
		  worksheet.getCell('I'+startrow).font = {name: 'Univers for KPMG',size: 10};
//      if(startrow!=10){
          worksheet.getCell('I'+startrow).dataValidation = {
              type: 'list',
              allowBlank: false,
              showErrorMessage: true,
              formulae: ['"OPEN, CLOSE"']
          };
          worksheet.getCell('I'+ startrow).value = "OPEN"
//      }          
      worksheet.getCell('J'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
      worksheet.getCell('J'+startrow).font = {name: 'Univers for KPMG',size: 10};
      worksheet.getCell('K'+startrow).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
      worksheet.getCell('K'+startrow).font = {name: 'Univers for KPMG',size: 10};
          worksheet.getCell('A'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('B'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('C'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('D'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('E'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('F'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('G'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('H'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('I'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('J'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          worksheet.getCell('K'+startrow).border = { top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
          if(worksheet.getCell('E'+startrow).value=="HIGH"){
            worksheet.getCell('E'+startrow).fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'FFFD3535'},bgColor:{argb:'FFFD3535'}};
          }
          else if(worksheet.getCell('E'+startrow).value=="MEDIUM"){
            worksheet.getCell('E'+startrow).fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'FFFFC000'},bgColor:{argb:'FFFFC000'}};
          }
          else if(worksheet.getCell('E'+startrow).value=="LOW"){
            worksheet.getCell('E'+startrow).fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'FF92D050'},bgColor:{argb:'FF92D050'}};
          }
          startrow++;

// code to insert images
          if(file_obj["obs"][file_obj["stock"][i]]["annexure"]){
			let youcantseeme = startrow-1

			worksheet2.mergeCells('B'+rowitr+':J'+rowitr);
            worksheet2.getCell('B'+ rowitr).value = file_obj["obs"][file_obj["stock"][i]]["observation"]
            worksheet2.getCell('B'+ rowitr).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
            worksheet2.getCell('B'+ rowitr).font = {name: 'Univers for KPMG',size: 10, color: { argb: 'FFFFFFFF' }};			
            worksheet2.getCell('B'+ rowitr).fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'FF00338D'},bgColor:{argb:'FF#00338D'}};
			worksheet.getCell('H'+youcantseeme).value = { text: 'Annexure', hyperlink: '#\'Annexures\'!B'+rowitr };
            rowitr+=3;
            let cases0 = Object.keys(file_obj["obs"][file_obj["stock"][i]]["images"])
			let cases =[]
			let thresholds = Object.keys(dimensions2["thresholds"])
			console.log("Here are the thresholds", thresholds);
			for(var j in cases0){
				cases.push(Number(cases0[j]))
			}
			cases.sort((a,b)=> a - b)
			if(cases.length>1){rowitr--;}
            for(var j in cases)
            {
              if(cases.length>1)
              {
                worksheet2.getCell('B'+rowitr).value = "Case " + cases[j];
				worksheet2.getCell('B'+ rowitr).font = {name: 'Univers for KPMG',size: 10, color: { argb: 'FF000000' }};			
                rowitr+=3;
              }
			  let image_keys=[]
              let image_keys0 = Object.keys(file_obj["obs"][file_obj["stock"][i]]["images"][cases[j]])
			  for(var k in image_keys0){
				  image_keys.push(Number(image_keys0[k]))
			  }
			  image_keys.sort((a,b)=> a - b)
			  var row_max=0;
              if(file_obj["img_limit"] == "infinite")
              {
                for(var k in image_keys)
                {
                   var img_path = file_obj["obs"][file_obj["stock"][i]]["images"][cases[j]][image_keys[k]]                   
				   // test code
				   var lol = sizeOf(img_path);
				   var selected_threshold
				   var current_min = 100
				   console.log("ratio is: ",lol.width/lol.height)
				   for(var z in thresholds){
					   if(current_min>Math.abs(Number(thresholds[z])-(lol.width/lol.height))){
						   current_min = Math.abs(Number(thresholds[z])-(lol.width/lol.height))
						   selected_threshold = thresholds[z];
    					   row_max = Math.max(row_max, dimensions2["thresholds"][thresholds[z]]["inc_row"])
					   }
				   }
				   console.log("selected threshold is: ", selected_threshold)
				   var dimen = convert(colitr)+String(rowitr)+':'+convert(colitr+dimensions2["thresholds"][selected_threshold]["inc_col"])+String(rowitr+dimensions2["thresholds"][selected_threshold]["inc_row"])
				   worksheet2.addImage(workbook.addImage({
                        filename: img_path,
                        extension: 'png',
                      }), dimen);
                   colitr= colitr+ dimensions2["thresholds"][selected_threshold]["inc_col"]+2
                 }
                 colitr=dimensions2["ini_col"]
                 rowitr= rowitr+row_max+3
              }
              else{
				if(image_keys.length % Number(file_obj["img_limit"]) === 0)
				{
					var l = Math.floor(image_keys.length/Number(file_obj["img_limit"]))
				}
				else{
					var l = Math.floor(image_keys.length/Number(file_obj["img_limit"])) +1
				}
                var m = Number(file_obj["img_limit"])
                var n = 0
                while(l--)
                {
   				row_max = 0;
                  while(m--)
                  {
					  console.log("ini_row_max", row_max)
                    if(n==image_keys.length-1)
                    {
                      var img_path = file_obj["obs"][file_obj["stock"][i]]["images"][cases[j]][image_keys[n]]
						var lol = sizeOf(img_path);
						var current_min = 100
						var selected_threshold;
						for(var z in thresholds){
							if(current_min>Math.abs(Number(thresholds[z])-(lol.width/lol.height))){
  							    current_min = Math.abs(Number(thresholds[z])-(lol.width/lol.height))
								selected_threshold = thresholds[z];
							}							
						}
						row_max = Math.max(row_max, dimensions2["thresholds"][selected_threshold]["inc_row"])
						console.log(row_max)						
                      var dimen = convert(colitr)+String(rowitr)+':'+convert(colitr+dimensions2["thresholds"][selected_threshold]["inc_col"])+String(rowitr+dimensions2["thresholds"][selected_threshold]["inc_row"])
						worksheet2.addImage(workbook.addImage({
                           filename: img_path,
                           extension: 'png',
                         }), dimen);
                      break;
                    }
                    var img_path = file_obj["obs"][file_obj["stock"][i]]["images"][cases[j]][image_keys[n]]
						var lol = sizeOf(img_path);
						var selected_threshold;
						var current_min = 100;
						for(var z in thresholds){
							if(current_min>Math.abs(Number(thresholds[z])-(lol.width/lol.height))){
								current_min = Math.abs(Number(thresholds[z])-(lol.width/lol.height))
								selected_threshold = thresholds[z];							
							}
						}
						row_max = Math.max(row_max, dimensions2["thresholds"][selected_threshold]["inc_row"])
						console.log(row_max)
                    var dimen = convert(colitr)+String(rowitr)+':'+convert(colitr+dimensions2["thresholds"][selected_threshold]["inc_col"])+String(rowitr+dimensions2["thresholds"][selected_threshold]["inc_row"])
                    worksheet2.addImage(workbook.addImage({
                           filename: img_path,
                           extension: 'png',
                      }), dimen);
                    colitr=colitr+dimensions2["thresholds"][selected_threshold]["inc_col"]+2
                    n++;
                  }
				  console.log("this is row_max: ", row_max)
				  m = Number(file_obj["img_limit"])
                  rowitr=rowitr+row_max+3
                  colitr=dimensions2["ini_col"]
                }
              }
            }
          }
          }
// code to update the table
        worksheet.getCell('C4').value = {formula: 'COUNTIF(E9:E10000, "HIGH")'}
        worksheet.getCell('C5').value = {formula: 'COUNTIF(E9:E10000, "MEDIUM")'}
        worksheet.getCell('C6').value = {formula: 'COUNTIF(E9:E10000, "LOW")'}
        worksheet.getCell('C7').value = {formula: 'C4+C5+C6'}

        worksheet.getCell('D4').value = {formula: 'COUNTIFS(E9:E10000,"HIGH",I9:I10000,"CLOSE")'}
        worksheet.getCell('D5').value = {formula: 'COUNTIFS(E9:E10000,"MEDIUM",I9:I10000,"CLOSE")'}
        worksheet.getCell('D6').value = {formula: 'COUNTIFS(E9:E10000,"LOW",I9:I10000,"CLOSE")'}
        worksheet.getCell('D7').value = {formula: 'D4+D5+D6'}

        worksheet.getCell('E4').value = {formula: 'COUNTIFS(E9:E10000,"HIGH",I9:I10000,"OPEN")'}
        worksheet.getCell('E5').value = {formula: 'COUNTIFS(E9:E10000,"MEDIUM",I9:I10000,"OPEN")'}
        worksheet.getCell('E6').value = {formula: 'COUNTIFS(E9:E10000,"LOW",I9:I10000,"OPEN")'}
        worksheet.getCell('E7').value = {formula: 'E4+E5+E6'}


           workbook.xlsx.writeFile(path.join('.','projects',filename.split(".")[0]+'.xlsx'))
            .then(function() {
              alert("REPORT CREATED. PLEASE FIND IN THE DIRECTORY:\n KREPORT\\projects")
          })
          .catch(err=> alert("FAILED. PREVIOUS REPORT MIGHT BE OPEN. CLOSE IT AND TRY AGAIN!"));
    });
}
