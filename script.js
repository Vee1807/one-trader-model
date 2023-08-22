// Document
let queryForm = document.getElementById("query-container");
let floatInput = document.getElementById("float-input")
let elements = queryForm.elements;
for (let i = 0; i < elements.length; i++) {
    elements[i].disabled = true;
}
// Query form (only float for now - no API yet)
queryForm.addEventListener('submit', e => {
    e.preventDefault()
    let floatValue = floatInput.value;
    modelTest(XL_row_object, floatValue);
})




// Source - https://stackoverflow.com/questions/8238407/how-to-parse-excel-xls-file-in-javascript-html5
let XL_row_object;
let ExcelToJSON = function() {  
    this.parseExcel = function(file) {
      let reader = new FileReader();
      reader.onload = function(e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, {
          type: 'binary'
        });
        workbook.SheetNames.forEach(function(sheetName) {
          // Here is your object
          XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          // let json_object = JSON.stringify(XL_row_object);
        })

        // Enable form
        for (let i = 0; i < elements.length; i++) {
            elements[i].disabled = false;
        }        
      };
  
      reader.onerror = function(ex) {
        console.log(ex);
      };
  
      reader.readAsBinaryString(file);
    };
    
  };


// Source - https://stackoverflow.com/questions/8238407/how-to-parse-excel-xls-file-in-javascript-html5
  function handleFileSelect(evt) {
    let files = evt.target.files; // FileList object
    let xl2json = new ExcelToJSON();
    xl2json.parseExcel(files[0]);
  }

// Upload button event listener
document.getElementById('upload').addEventListener('change', handleFileSelect, false);


// Model
function modelTest(data_obj, floatValue) {
    let j = 12;
    let row = [];
    let table = [];
    let nrow = data_obj.length;
    console.log(floatValue)
    let loopMax = nrow * 2 - 1;

    // First 2 rows
    row.push(-1 * data_obj[0].Volume);
    table.push(row);
    row = [];
    row.push( 1 * table[0][0] +  1 * floatValue);
    row.push(- 1 * table[0][0]);
    table.push(row);


    let colCount = 2;
    let colCount3 = 1;
    let index = 1;

    // Loop through rows
    for(let i = 1; i < loopMax; i++) {
        colCount3 = 1;
        row = [];
        f = 0;

      // Loop through columns
        for(let l = 0; l < colCount; l++) {
            if (i % 2 == 0) {
                row.push(1 * table[i][l] + 1 * table[i - 1][l]);
                f = 1;
            } else {
            //row.push(`=R${j}C3*(R[-1]C[0]/R5C5)*-1`);
                row.push(- 1 * data_obj[index].Volume * table[i][l] / floatValue);
          
          
        }
        colCount3++;
      }
        if (f == 1) {
            row.push(1 * data_obj[index].Volume)
            colCount++;
            j += 2;
      }/* else {
        row.push('=""');
        
      }*/
  
      /*for(let k = 0; k < ncol - colCount3; k++) {
        row.push('=""');
      }     */

        table.push(row)

        // Increment i every 2 loops
        if (i % 2 == 0) {
        index++;
        }
    }
    

    // Second table
    let prices = [];

    // Convert xx.x% (String) to xx.x (Number)
    let changeStr = data_obj[nrow - 1]['%Chg'];
    changeStr = changeStr.slice(0, -1);
    let change = changeStr * 1
    let finalValues = [];
    let sum = 0
    for (let i = 0; i < nrow; i++) {
        prices.push(1 * data_obj[i].Last);
        finalValues.push(- 1 * prices[i] * table[table.length - 1][i + 1] * change / 100);

        // Calculate sum
        sum += finalValues[i];
    }
    console.log(sum)
  }



