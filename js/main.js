


const excel_file = document.getElementById("excel_file");

const excel_left = document.getElementById("excel_left");
const excel_right = document.getElementById("excel_right");

const CgpaId = document.getElementById("cgpa_info");
const excel_additional = document.getElementById("excel_additional");

excel_file.addEventListener("change", (event) => {
  if (
    ![
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
    ].includes(event.target.files[0].type)
  ) {
    document.getElementById("excel_data").innerHTML =
      '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';

    excel_file.value = "";

    return false;
  }
  var reader = new FileReader();

  reader.readAsArrayBuffer(event.target.files[0]);

  reader.onload = function (event) {
    var data = new Uint8Array(reader.result);

    var work_book = XLSX.read(data, { type: "array" });

    var sheet_name = work_book.SheetNames;

    var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {
      header: 1,
    });




    /* Roll No*/
    const RollNo = sheet_data[1][0];
    // RollNo = "<p>RollNo.  </p>" + RollNo;
    // document.getElementById("excel_data").innerHTML =RollNo;

    /*Name*/
    const Name = sheet_data[1][1];

    /*Year*/
    var Year = RollNo.slice(2, 4);
    var graduationYear = parseInt(Year) + 4;

    var student_info = "<span><b> Name&nbsp&nbsp;&nbsp;&emsp;:</b> &nbsp;" + Name + "</span><br>";
    student_info += "<span><b>    Roll&nbspNo&emsp;:</b> &nbsp;" + RollNo + "</span>";

    var student_year = "<p>&emsp;&emsp;&emsp;&emsp;<b>Year of Study&emsp;&nbsp</b> " + ":20" + Year + "-" + "20" + graduationYear + "</p>";

    document.getElementById("student_details").innerHTML = student_info;
    document.getElementById("student_year_details").innerHTML = student_year;

    /*calculating credits*/
    var TotalCredits = 0;
    var nonAdditionalCredits = 0;
    for (var row = 1; row < sheet_data.length; row++) {
      TotalCredits += parseFloat(sheet_data[row][9]);
    }


    /*------------------------------------calculating cgpa-------------------------*/
    const Hashmap = new Map([
      ['A', 10],
      ['A-', 9],
      ['B', 8],
      ['B-', 7],
      ['C', 6],
      ['C-', 5],
      ['D', 4],
      ['AU', 0],
      ['S', 0],
      ['F', 0],
      ['U', 0],
      ['', 0]
    ]);

    /*---------------------------grading scheme table-------------------------------------*/
    // grading_table = '<table class="table table-bordered">';
    // grading_table +=  '<tr> <th> <b>Grades</b></th> </tr>';
    // grading_table += '<tr><th> <b>Points</b></th></tr>';

    // var it = 0;
    // for(var key in Hashmap){
    //   if(it == Hashmap.length-1){
    //     break;
    //   }
    //   grading_table += '<tr> <td>' + key + '</td></tr>';
    //   grading_table += '<tr><td>' + Hashmap[key] + '</td></tr>';
    // }

    // grading_table += '</table>';
    // document.getElementById("grading_table").innerHTML = grading_table;

    var Additional_Course = [];

    var sum = 0;
    for (var row = 1; row < sheet_data.length; row++) {
      var credit = parseInt(sheet_data[row][6]);
      // console.log("credit : " + credit);
      var grade = sheet_data[row][10];
      // console.log("grade: " + grade);
      var gradeValue = Hashmap.get(grade);
      // console.log("GradeValue: " + gradeValue);
      if (sheet_data[row][7] != "Additional" ) {  // not including additional grade
        sum += credit * gradeValue;
        nonAdditionalCredits += parseFloat(sheet_data[row][6]);
        // console.log("sum" + sum + "--> row : " + row);
      }else{
        Additional_Course.push(sheet_data[row][4]);
      }

    }

    console.log(Additional_Course);
    var cgpa = (sum / nonAdditionalCredits).toFixed(2);
    console.log(cgpa);

    console.log("TotalCredits -> " + TotalCredits);
    console.log("nonAdditionalCredits-> " + nonAdditionalCredits);
    /*----------------------------------------------------------------*/

    /*creating a function to check wether the required header is present or not*/
    var check = [0,1,2,3,7,8,9,11]
    function not(dat, arr) {
      for (var i = 0; i < arr.length; i++) {
        if (arr[i] == dat) {
          return false;
        }
      }
      return true;

    }

    /*--------------------hash map for converting the header names as required for particular column--------------------------------------------*/
    const HeaderMap = new Map([
      [4, "Course No."],
      [5, "Course Name"],
      [6 , "Credits"],
      [10, "Grade"]
    ]);

    /*----------------------------------------------------------------*/

    /*creating tables*/

    var half = Math.floor(sheet_data.length / 2);
    var col_length = sheet_data[0].length;
    console.log(col_length);

    /* sheet data lenght odd length case */
    if (sheet_data.length % 2 != 0) {
      /*-------------left table------------------*/
      var table_output_left = '<table class="table table-bordered">';

      for (var row = 0; row <= half; row++) {
        table_output_left += "<tr>";
        console.log(row);

        for (var cell = 0; cell < col_length; cell++) {
          
          if (not(cell, check)) {
            if (row == 0) {
              var headerValue = HeaderMap.get(cell);
              table_output_left += "<th style=text-align: center>" + headerValue + "</th>";

              // if (row == 0) {
              //   if(cell==4){
              //     table_output_right += "<th>Course No.</th>";
              //   }
              //   else if(cell==9){
              //     table_output_right += "<th>Credits</th>";
              //   }else{
              //     table_output_right += "<th>" + sheet_data[0][cell] + "</th>";
              //   }
            } else {
              if (cell != 5) {
              table_output_left += "<td style=text-align:center>" + sheet_data[row][cell] + "</td>";

              }
              else {
                table_output_left += "<td>" + sheet_data[row][cell] + "</td>";

              }
            }

          }

        }
        table_output_left += "</tr>";
      }
      table_output_left += "</table>";
      document.getElementById("excel_left").innerHTML = table_output_left;







      /*-------------right table -----------------*/
      var table_output_right = '<table class="table table-bordered">';

      for (var row = half; row < sheet_data.length; row++) {
        table_output_right += "<tr>";
        console.log(row);

        for (var cell = 0; cell < col_length; cell++) {
          if (not(cell, check)) {
            if (row == half) {
              var headerValue = HeaderMap.get(cell);
              table_output_right += "<th style=text-align: center>" + headerValue + "</th>";
              // if(cell==4){
              //   table_output_right += "<th>" + "Course No." + "</th>";
              // }
              // else if(cell==9){
              //   table_output_right += "<th>" + "Credits" + "</th>";
              // }else{
              //   table_output_right += "<th>" + sheet_data[0][cell] + "</th>";
              // }
            } else {
              if (cell != 5) {
                table_output_right += "<td style=text-align:center>" + sheet_data[row][cell] + "</td>";

              }
              else {
                table_output_right += "<td>" + sheet_data[row][cell] + "</td>";

              }
            }
          }
        }

        table_output_right += "</tr>";
      }

      table_output_right += "</table>";
      document.getElementById("excel_right").innerHTML = table_output_right;

    }
    else {
      /* shett data length even length case*/

      /*-------------left table------------------*/
      var table_output_left = '<table class="table table-bordered">';

      for (var row = 0; row <= half; row++) {
        table_output_left += "<tr>";
        console.log(row);

        for (var cell = 0; cell < col_length; cell++) {
          if (not(cell, check)) {
            if (row == 0) {
              var headerValue = HeaderMap.get(cell);
              table_output_left += "<th style=text-align: center>" + headerValue + "</th>";
              // if(cell==4){
              //   table_output_right += "<th>" + "Course No." + "</th>";
              // }
              // else if(cell==9){
              //   table_output_right += "<th>" + "Credits" + "</th>";
              // }else{
              //   table_output_right += "<th>" + sheet_data[0][cell] + "</th>";
              // }
            } else {
              if (cell != 5) {
                table_output_left += "<td style=text-align:center>" + sheet_data[row][cell] + "</td>";

              }
              else {
                table_output_left += "<td>" + sheet_data[row][cell] + "</td>";

              }
            }
          }
        }
        table_output_left += "</tr>";
      }

      table_output_left += "</table>";
      document.getElementById("excel_left").innerHTML = table_output_left;


      /*-------------right table -----------------*/
      var table_output_right = '<table class="table table-bordered">';

      for (var row = half; row <= sheet_data.length; row++) {
        table_output_right += "<tr>";
        // console.log(row);
        for (var cell = 0; cell < col_length; cell++) {
          if (not(cell, check)) {
            if (row == half) {
              var headerValue = HeaderMap.get(cell);
              table_output_right += "<th style=text-align: center>" + headerValue + "</th>"
            } else {
              if (cell != 5) {
                table_output_right += "<td style=text-align:center>" + sheet_data[row][cell] + "</td>";

              }
              else {
                table_output_right += "<td>" + sheet_data[row][cell] + "</td>";

              }
            }
          }
        }
        table_output_right += "</tr>";
      }

      table_output_right += "</table>";
      document.getElementById("excel_right").innerHTML = table_output_right;
    }






    /**cgpa info**/
    var cgpa_cal = "<b><h5 >Cumulative Grade Point Average<small>(Out of 10.00)&nbsp:&nbsp </small>" + cgpa + "</b> </h5>";
    cgpa_cal += "<h5> <b>Total Credits Earned</b> &emsp; &emsp;&emsp; &emsp; &emsp;&emsp; &emsp; &emsp;&emsp; : &nbsp<b>" + TotalCredits + "</b></h5>"
    document.getElementById("cgpa_info").innerHTML = cgpa_cal;

    /********************************/

    /** Additional courses table */
    // document.getElementById("additonal_heading").innerHTML = '<table class="table table-bordered"> <tr> <th> Additional Courses </th> </tr>  </table>';
    table_additional = '<table class="table table-bordered" style ="margin-bottom: 0px; text-align: center; width:100%">';
    table_additional += "<tr> <th> <p>Additional Courses</p> </th> </tr>"
    table_additional += "</table>";
    document.getElementById("heading").innerHTML = table_additional;



    table_additional = '<table class="table table-bordered " style="width:100%">';
    table_additional += "<tr> "
    table_additional += "<th>" + "Course No" + "</th>";
    table_additional += "<th>" + "Course Name" + "</th>";
    table_additional += "<th>" + "Credits" + "</th>";
    table_additional += "<th>" + "Grade" + "</th>";
    table_additional += "</tr>";
    for (var i = 0; i < Additional_Course.length; i++) {
      table_additional += "<tr>";
      for (var j = 0; j < sheet_data.length; j++) {
        if (sheet_data[j][4] == Additional_Course[i]) {
          table_additional += "<td>" + Additional_Course[i] + "</td>";
          table_additional += "<td>" + sheet_data[j][5] + "</td>";
          table_additional += "<td>" + sheet_data[j][6] + "</td>";
          table_additional += "<td>" + sheet_data[j][10] + "</td>";

        }
      }
      table_additional += "</tr>";

    }
    table_additional += "</table>";
    document.getElementById("excel_additional").innerHTML = table_additional;

    excel_file.value = "";


  };
});

// var div = document.getElementById("resultContainer");
// // get reference to button
// var btn = document.getElementById("DownloadBtn");
// // add event listener for the button, for action "click"
// btn.addEventListener("click", PrintElem(div));
// function PrintElem(elem) {
//   Popup($(elem).html());
// }


function Popup(data) {
  var mywindow = window.open('', 'new div', 'height=400,width=600');
  mywindow.document.write('<html><head><title></title>');
  mywindow.document.write('<link rel="stylesheet" href="css/style.css" type="text/css" />');
  mywindow.document.write('</head><body >');
  mywindow.document.write(data);
  mywindow.document.write('</body></html>');
  mywindow.document.close();
  mywindow.focus();
  setTimeout(function(){mywindow.print();},500);
  mywindow.close();

  return true;
}




// document.getElementById("getPDF").addEventListener("click", function(){ 
//   var newWindowContent = document.getElementById('resultContainer').innerHTML;
//   var newWindow = window.open("", "", "width=500,height=400");
//   newWindow.document.write(newWindowContent);
// });

// document.getElementById("getPDF").addEventListener("click", function (){
//     var divToPrint = document.getElementById('resultContainer');
//     var popupWin = window.open('', '_blank', 'width=300,height=300');
//     popupWin.document.open();
//     popupWin.document.write('<html><link href="css/style.css"  rel="stylesheet" /><body onload="window.print()">' + divToPrint.innerHTML + '</html>');
//     popupWin.document.close();
// });


function getPDF() {

  var HTML_Width = $(".resultContainer").width() + 4;
  var HTML_Height = $(".resultContainer").height() + 4;
  var top_left_margin = 60;
  var PDF_Width = HTML_Width + (top_left_margin * 2);
  var PDF_Height = (PDF_Width * 1.5) + (top_left_margin * 2);
  var canvas_image_width = HTML_Width;
  var canvas_image_height = HTML_Height;

  var totalPDFPages = Math.ceil(HTML_Height / PDF_Height) - 1;


  html2canvas($(".resultContainer")[0], { allowTaint: true }).then(function (canvas) {
    canvas.getContext('2d');

    console.log(canvas.height + "  " + canvas.width);


    var imgData = canvas.toDataURL("image/jpeg", 1.0);
    var pdf = new jsPDF('p', 'pt', [PDF_Width, PDF_Height]);
    pdf.addImage(imgData, 'JPG', top_left_margin, top_left_margin, canvas_image_width, canvas_image_height);


    for (var i = 1; i <= totalPDFPages; i++) {
      pdf.addPage(PDF_Width, PDF_Height);
      pdf.addImage(imgData, 'JPG', top_left_margin, -(PDF_Height * i) + (top_left_margin * 4), canvas_image_width, canvas_image_height);
    }

    pdf.save("HTML-Document.pdf");
  });
};
