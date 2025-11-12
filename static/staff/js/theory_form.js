      let totalStudents = 0;
      let passedStudents = 0;

      function generateTable() {
        const numStudents = parseInt(
          document.getElementById("numStudents").value
        );
        if (isNaN(numStudents) || numStudents <= 0) {
          alert("Please enter a valid number of students.");
          return;
        }
        const table = document.getElementById("studentsTable");
        if (!table) {
          console.error("Table element with ID 'studentsTable' not found.");
          return;
        }

        const tbody = table.querySelector("tbody");
        if (!tbody) {
          console.error("Table body not found.");
          return;
        }

        document.getElementById("all").textContent = numStudents;
        tbody.innerHTML = ""; // Clear previous rows

        let rows = "";
        for (let i = 1; i <= numStudents; i++) {
          rows += `
            <tr>
                <td><input type="number" id="reg-${i}" class="form-control" placeholder="Enter Reg. No"></td>
                <td><input type="text" id="name-${i}" class="form-control" placeholder="Enter Name"></td>
                <td><input type="text" id="ese-${i}" class="form-control" oninput="calculateTotal1(${i});grade(${i});over()"></td>
                <td><input type="text" id="cia-${i}" class="form-control" oninput="calculateTotal1(${i});CIA(${i});over()"></td>
                <td><input type="text" id="total-${i}" class="form-control" readonly></td>
            </tr>`;
        }

        tbody.insertAdjacentHTML("beforeend", rows); // Append all rows at once

        // Arrow key navigation
        setTimeout(() => {
          const inputs = tbody.querySelectorAll("input");
          inputs.forEach((input, index) => {
            input.addEventListener("keydown", (e) => {
              if (
                ["ArrowRight", "ArrowLeft", "ArrowDown", "ArrowUp"].includes(
                  e.key
                )
              ) {
                e.preventDefault();
                let rowLength = 5;
                let newIndex = index;
                switch (e.key) {
                  case "ArrowRight":
                    newIndex = index + 1;
                    break;
                  case "ArrowLeft":
                    newIndex = index - 1;
                    break;
                  case "ArrowDown":
                    newIndex = index + rowLength;
                    break;
                  case "ArrowUp":
                    newIndex = index - rowLength;
                    break;
                }
                if (inputs[newIndex]) {
                  inputs[newIndex].focus();
                }
              }
            });
          });
        }, 0);
      }

      function handleFile() {
        const fileInput = document.getElementById("excelFile");
        const file = fileInput.files[0];

        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (event) {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(sheet, {
            header: "A",
            raw: false,
          });

          // Fill programme details
          document.getElementById("programme").value = sheet["A2"]
            ? sheet["A2"].v.replace("Programme: ", "")
            : ""; // Adjust cell reference
          document.getElementById("courseName").value = sheet["A3"]
            ? sheet["A3"].v.replace("Course Name: ", "")
            : "";
          document.getElementById("courseCode").value = sheet["A4"]
            ? sheet["A4"].v.replace("Course Code: ", "")
            : "";
          document.getElementById("academicYear").value = sheet["F2"]
            ? sheet["F2"].v
            : "";
          document.getElementById("semester").value = sheet["F3"]
            ? sheet["F3"].v
            : "";

          // Get number of students from cell I17 (adjust this as per your Excel format)
          const numStudents = parseInt(sheet["I17"] ? sheet["I17"].v : 0);
          document.getElementById("numStudents").value = numStudents;
          generateTable();

          // Fill student data (starting from row 7)
          for (let i = 0; i < numStudents; i++) {
            const regNo = sheet[`B${7 + i}`] ? sheet[`B${7 + i}`].v : "";
            const name = sheet[`C${7 + i}`] ? sheet[`C${7 + i}`].v : "";
            const ese = sheet[`D${7 + i}`] ? sheet[`D${7 + i}`].v : "";
            const cia = sheet[`E${7 + i}`] ? sheet[`E${7 + i}`].v : "";

            document.getElementById(`reg-${i + 1}`).value = regNo;
            document.getElementById(`name-${i + 1}`).value = name;
            document.getElementById(`ese-${i + 1}`).value = ese;
            document.getElementById(`cia-${i + 1}`).value = cia;

            // Manually trigger calculation functions
            calculateTotal1(i + 1);
            grade(i + 1);
            over();
            CIA(i + 1);
          }

          // Fill Course Outcome Survey (COs) table (adjust cell references)
          for (let i = 1; i <= 5; i++) {
            document.getElementById(`Excellent${i}`).value = sheet[`I${i + 30}`]
              ? sheet[`I${i + 30}`].v
              : 0;
            document.getElementById(`Good${i}`).value = sheet[`J${i + 30}`]
              ? sheet[`J${i + 30}`].v
              : 0;
            document.getElementById(`Fair${i}`).value = sheet[`K${i + 30}`]
              ? sheet[`K${i + 30}`].v
              : 0;
            COs();
            over();
            updateStatistics();
          }
        };
        reader.readAsArrayBuffer(file);
      }

      function calculateTotal1(row) {
        const ese =
          parseFloat(document.getElementById(`ese-${row}`).value) || 0;
        const cia =
          parseFloat(document.getElementById(`cia-${row}`).value) || 0;

        document.getElementById(`total-${row}`).value = ese + cia;
      }
      let ciacount = 0;
      const processedRows = new Set(); // To track rows already counted

      function CIA(row) {
        const numStudents =
          parseInt(document.getElementById("numStudents").value) || 0;
        const cia =
          parseFloat(document.getElementById(`cia-${row}`).value) || 0;
        const ciaThreshold = 17.5;

        if (cia > ciaThreshold) {
          if (!processedRows.has(row)) {
            // Only increment if the row is not already processed
            ciacount += 1;
            processedRows.add(row); // Mark this row as processed
          }
        } else {
          if (processedRows.has(row)) {
            // If row is processed but now fails the threshold
            ciacount -= 1;
            processedRows.delete(row); // Remove this row from the processed set
          }
        }

        // percentage
        document.getElementById("cia70").textContent = ciacount;

        const ciaper = ((ciacount / numStudents) * 100).toFixed(2);
        document.getElementById("ciaper").textContent = ciaper + "%";

        let ciauni;
        if (ciaper < 60) {
          ciauni = 0;
          document.getElementById("ciauni").textContent = "0";
        } else if (ciaper < 70) {
          ciauni = 1;
          document.getElementById("ciauni").textContent = "1";
        } else if (ciaper < 80) {
          ciauni = 2;
          document.getElementById("ciauni").textContent = "2";
        } else {
          ciauni = 3;
          document.getElementById("ciauni").textContent = "3";
        }
        document.getElementById("IA1").textContent = (ciauni * 0.4).toFixed(2);
        document.getElementById("IA2").textContent = (ciauni * 0.4).toFixed(2);
        document.getElementById("IA3").textContent = (ciauni * 0.4).toFixed(2);
        document.getElementById("IA4").textContent = (ciauni * 0.4).toFixed(2);
        document.getElementById("IA5").textContent = (ciauni * 0.4).toFixed(2);
      }

      let previousGrades = {};
      let gradeCounts = { O: 0, D: 0, Aplus: 0, A: 0, B: 0, C: 0, U: 0, AA: 0 };

      function grade(row) {
        const eseInput = document.getElementById(`ese-${row}`);
        const eseValue = parseFloat(eseInput.value) || 0;
        let grade = "";

        // Grade calculation logic
        if (eseInput.value.toLowerCase() === "aa") {
          grade = "AA";
        } else if (eseValue >= 65 && eseValue <= 75) {
          grade = "O";
        } else if (eseValue >= 55 && eseValue <= 64) {
          grade = "D";
        } else if (eseValue >= 50 && eseValue <= 54) {
          grade = "Aplus";
        } else if (eseValue >= 45 && eseValue <= 49) {
          grade = "A";
        } else if (eseValue >= 36 && eseValue <= 44) {
          grade = "B";
        } else if (eseValue >= 30 && eseValue <= 35) {
          grade = "C";
        } else if (eseValue >= 0 && eseValue <= 29) {
          grade = "U";
        }

        // Update gradeCounts
        if (previousGrades[row]) {
          gradeCounts[previousGrades[row]]--;
        }
        gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;
        previousGrades[row] = grade;

        // Update UI for grade counts
        updateGradeUI();

        // Calculate and update percentages
        updateStatistics();
      }

      function updateGradeUI() {
        document.getElementById("O").textContent = gradeCounts.O || 0;
        document.getElementById("D").textContent = gradeCounts.D || 0;
        document.getElementById("A+").textContent = gradeCounts.Aplus || 0;
        document.getElementById("A").textContent = gradeCounts.A || 0;
        document.getElementById("B").textContent = gradeCounts.B || 0;
        document.getElementById("C").textContent = gradeCounts.C || 0;
        document.getElementById("U").textContent = gradeCounts.U || 0;
        document.getElementById("AAA").textContent = gradeCounts.AA || 0;
      }

      function updateStatistics() {
        const totalStudents = Object.keys(previousGrades).length;
        const passedStudents =
          gradeCounts.O +
          gradeCounts.D +
          gradeCounts.Aplus +
          gradeCounts.A +
          gradeCounts.B +
          gradeCounts.C;
        const percentage = ((passedStudents / totalStudents) * 100).toFixed(2);

        document.getElementById("pass").textContent = passedStudents || 0;
        document.getElementById("studentsper").textContent = percentage + "%";

        let universityPoints = 0;
        if (percentage < 50) universityPoints = 0;
        else if (percentage < 60) universityPoints = 1;
        else if (percentage < 70) universityPoints = 2;
        else universityPoints = 3;
        document.getElementById("university").textContent = universityPoints;
        for (let i = 1; i <= 5; i++) {
          document.getElementById(`point2${i}`).textContent = Math.round(
            universityPoints * 0.6
          );
        }
      }

      function COs() {
        const Excellent1 = document.getElementById("Excellent1");
        const Excellent2 = document.getElementById("Excellent2");
        const Excellent3 = document.getElementById("Excellent3");
        const Excellent4 = document.getElementById("Excellent4");
        const Excellent5 = document.getElementById("Excellent5");
        const Good1 = document.getElementById("Good1");
        const Good2 = document.getElementById("Good2");
        const Good3 = document.getElementById("Good3");
        const Good4 = document.getElementById("Good4");
        const Good5 = document.getElementById("Good5");
        const Fair1 = document.getElementById("Fair1");
        const Fair2 = document.getElementById("Fair2");
        const Fair3 = document.getElementById("Fair3");
        const Fair4 = document.getElementById("Fair4");
        const Fair5 = document.getElementById("Fair5");
        const Total1 = document.getElementById("Total1");
        const Total2 = document.getElementById("Total2");
        const Total3 = document.getElementById("Total3");
        const Total4 = document.getElementById("Total4");
        const Total5 = document.getElementById("Total5");
        const Persentage1 = document.getElementById("Persentage1");
        const Persentage2 = document.getElementById("Persentage2");
        const Persentage3 = document.getElementById("Persentage3");
        const Persentage4 = document.getElementById("Persentage4");
        const Persentage5 = document.getElementById("Persentage5");
        const Outcome1 = document.getElementById("Outcome1");
        const Outcome2 = document.getElementById("Outcome2");
        const Outcome3 = document.getElementById("Outcome3");
        const Outcome4 = document.getElementById("Outcome4");
        const Outcome5 = document.getElementById("Outcome5");

        const ex1 = parseFloat(Excellent1.value) || 0;
        const go1 = parseFloat(Good1.value) || 0;
        const fa1 = parseFloat(Fair1.value) || 0;
        const ex2 = parseFloat(Excellent2.value) || 0;
        const go2 = parseFloat(Good2.value) || 0;
        const fa2 = parseFloat(Fair2.value) || 0;
        const ex3 = parseFloat(Excellent3.value) || 0;
        const go3 = parseFloat(Good3.value) || 0;
        const fa3 = parseFloat(Fair3.value) || 0;
        const ex4 = parseFloat(Excellent4.value) || 0;
        const go4 = parseFloat(Good4.value) || 0;
        const fa4 = parseFloat(Fair4.value) || 0;
        const ex5 = parseFloat(Excellent5.value) || 0;
        const go5 = parseFloat(Good5.value) || 0;
        const fa5 = parseFloat(Fair5.value) || 0;
        let to1 = 0;
        let to2 = 0;
        let to3 = 0;
        let to4 = 0;
        let to5 = 0;
        let oc1 = 0;
        let oc2 = 0;
        let oc3 = 0;
        let oc4 = 0;
        let oc5 = 0;
        to1 = ex1 + go1 + fa1;
        Total1.textContent = to1;
        if (to1 > 0) {
          oc1 = parseFloat(((ex1 / to1) * 100).toFixed(2));
          Persentage1.textContent = oc1 + "%";
        } else Persentage1.textContent = "0.00";
        to2 = ex2 + go2 + fa2;
        Total2.textContent = to2;
        if (to2 > 0) {
          oc2 = parseFloat(((ex2 / to2) * 100).toFixed(2));
          Persentage2.textContent = oc2 + "%";
        } else Persentage2.textContent = "0.00";
        to3 = ex3 + go3 + fa3;
        Total3.textContent = to3;
        if (to3 > 0) {
          oc3 = parseFloat(((ex3 / to3) * 100).toFixed(2));
          Persentage3.textContent = oc3 + "%";
        } else Persentage3.textContent = "0.00";
        to4 = ex4 + go4 + fa4;
        Total4.textContent = to4;
        if (to4 > 0) {
          oc4 = parseFloat(((ex4 / to4) * 100).toFixed(2));
          Persentage4.textContent = oc4 + "%";
        } else Persentage4.textContent = "0.00";
        to5 = ex5 + go5 + fa5;
        Total5.textContent = to5;
        if (to5 > 0) {
          oc5 = parseFloat(((ex5 / to5) * 100).toFixed(2));
          Persentage5.textContent = oc5 + "%";
        } else Persentage5.textContent = "0.00";

        if (oc1 < 61) {
          Outcome1.textContent = "0";
        } else if (oc1 < 71) {
          Outcome1.textContent = "1";
        } else if (oc1 < 81) {
          Outcome1.textContent = "2";
        } else {
          Outcome1.textContent = "3";
        }
        if (oc2 < 61) {
          Outcome2.textContent = "0";
        } else if (oc2 < 71) {
          Outcome2.textContent = "1";
        } else if (oc2 < 81) {
          Outcome2.textContent = "2";
        } else {
          Outcome2.textContent = "3";
        }
        if (oc3 < 61) {
          Outcome3.textContent = "0";
        } else if (oc3 < 71) {
          Outcome3.textContent = "1";
        } else if (oc3 < 81) {
          Outcome3.textContent = "2";
        } else {
          Outcome3.textContent = "3";
        }
        if (oc4 < 61) {
          Outcome4.textContent = "0";
        } else if (oc4 < 71) {
          Outcome4.textContent = "1";
        } else if (oc4 < 81) {
          Outcome4.textContent = "2";
        } else {
          Outcome4.textContent = "3";
        }
        if (oc5 < 61) {
          Outcome5.textContent = "0";
        } else if (oc5 < 71) {
          Outcome5.textContent = "1";
        } else if (oc5 < 81) {
          Outcome5.textContent = "2";
        } else {
          Outcome5.textContent = "3";
        }
        const ida1 = (document.getElementById("ida1").textContent =
          Outcome1.textContent);
        const ida2 = (document.getElementById("ida2").textContent =
          Outcome2.textContent);
        const ida3 = (document.getElementById("ida3").textContent =
          Outcome3.textContent);
        const ida4 = (document.getElementById("ida4").textContent =
          Outcome4.textContent);
        const ida5 = (document.getElementById("ida5").textContent =
          Outcome5.textContent);

        //IDA
        document.getElementById("point1").textContent = (
          Outcome1.textContent * 0.2
        ).toFixed(2);
        document.getElementById("point2").textContent = (
          Outcome2.textContent * 0.2
        ).toFixed(2);
        document.getElementById("point3").textContent = (
          Outcome3.textContent * 0.2
        ).toFixed(2);
        document.getElementById("point4").textContent = (
          Outcome4.textContent * 0.2
        ).toFixed(2);
        document.getElementById("point5").textContent = (
          Outcome5.textContent * 0.2
        ).toFixed(2);
        document.getElementById("allpoint").textContent = Math.round(
          (parseInt(Outcome1.textContent) +
            parseInt(Outcome2.textContent) +
            parseInt(Outcome3.textContent) +
            parseInt(Outcome4.textContent) +
            parseInt(Outcome5.textContent)) /
            5
        );
      }

      function over() {
        //DA
        const point11 =
          parseFloat(document.getElementById("point1").textContent) || 0;
        const point12 =
          parseFloat(document.getElementById("point2").textContent) || 0;
        const point13 =
          parseFloat(document.getElementById("point3").textContent) || 0;
        const point14 =
          parseFloat(document.getElementById("point4").textContent) || 0;
        const point15 =
          parseFloat(document.getElementById("point5").textContent) || 0;

        const IA1 = parseInt(document.getElementById("IA1").textContent) || 0;
        const IA2 = parseInt(document.getElementById("IA2").textContent) || 0;
        const IA3 = parseInt(document.getElementById("IA3").textContent) || 0;
        const IA4 = parseInt(document.getElementById("IA4").textContent) || 0;
        const IA5 = parseInt(document.getElementById("IA5").textContent) || 0;

        const point21 =
          parseInt(document.getElementById("point21").textContent) || 0;
        const point22 =
          parseInt(document.getElementById("point22").textContent) || 0;
        const point23 =
          parseInt(document.getElementById("point23").textContent) || 0;
        const point24 =
          parseInt(document.getElementById("point24").textContent) || 0;
        const point25 =
          parseInt(document.getElementById("point25").textContent) || 0;

        const point01 = Math.round(IA1 + point21);
        const point02 = Math.round(IA2 + point22);
        const point03 = Math.round(IA3 + point23);
        const point04 = Math.round(IA4 + point24);
        const point05 = Math.round(IA5 + point25);

        document.getElementById("totalPoint1").textContent = point01;
        document.getElementById("totalPoint2").textContent = point02;
        document.getElementById("totalPoint3").textContent = point03;
        document.getElementById("totalPoint4").textContent = point04;
        document.getElementById("totalPoint5").textContent = point05;

        const average = Math.round(
          (point01 + point02 + point03 + point04 + point05) / 5
        );
        document.getElementById("avarage").textContent = average;

        const max1 = Math.round(point01 * 0.8 + point11);
        const max2 = Math.round(point02 * 0.8 + point12);
        const max3 = Math.round(point03 * 0.8 + point13);
        const max4 = Math.round(point04 * 0.8 + point14);
        const max5 = Math.round(point05 * 0.8 + point15);
        //over all attainment
        document.getElementById("max1").textContent = max1;
        document.getElementById("max2").textContent = max2;
        document.getElementById("max3").textContent = max3;
        document.getElementById("max4").textContent = max4;
        document.getElementById("max5").textContent = max5;

        // over all attainment total
        const add1 = parseInt(document.getElementById("max1").textContent) || 0;
        const add2 = parseInt(document.getElementById("max2").textContent) || 0;
        const add3 = parseInt(document.getElementById("max3").textContent) || 0;
        const add4 = parseInt(document.getElementById("max4").textContent) || 0;
        const add5 = parseInt(document.getElementById("max5").textContent) || 0;

        //add all
        document.getElementById("alltotal").textContent =
          add1 + add2 + add3 + add4 + add5;

        //ultramax
        const ultra =
          parseInt(document.getElementById("alltotal").textContent) || 0;
        document.getElementById("ultramax").textContent = Math.round(ultra / 5);
      }

      //post data
      function collectData() {
        // Collect Header Information
        const headerData = {
          programme: document.getElementById("programme").value,
          courseName: document.getElementById("courseName").value,
          courseCode: document.getElementById("courseCode").value,
          academicYear: document.getElementById("academicYear").value,
          semester: document.getElementById("semester").value,
        };

        // Collect Students Table Data

        const students = [];
        const rows = document
          .getElementById("studentsTable")
          .querySelectorAll("tbody tr");
        rows.forEach((row, index) => {
          students.push({
            regNo: document.getElementById(`reg-${index + 1}`).value,
            name: document.getElementById(`name-${index + 1}`).value,
            ese:
              parseFloat(document.getElementById(`ese-${index + 1}`).value) ||
              0,
            cia:
              parseFloat(document.getElementById(`cia-${index + 1}`).value) ||
              0,
            total:
              parseFloat(document.getElementById(`total-${index + 1}`).value) ||
              0,
          });
        });
        // Collect CO Survey Data
        const coSurveyData = Array.from({ length: 5 }, (_, i) => ({
          co: `CO${i + 1}`,
          excellent: document.getElementById(`Excellent${i + 1}`).value || 0,
          good: document.getElementById(`Good${i + 1}`).value || 0,
          fair: document.getElementById(`Fair${i + 1}`).value || 0,
          total: document.getElementById(`Total${i + 1}`).textContent || 0,
          percentage:
            document.getElementById(`Persentage${i + 1}`).textContent || 0,
          outcome: document.getElementById(`Outcome${i + 1}`).textContent || 0,
        }));

        // Collect ESE, CIA, and Overall Data
        const eseData = {
          O: document.getElementById("O").textContent || 0,
          D: document.getElementById("D").textContent || 0,
          "A+": document.getElementById("A+").textContent || 0,
          A: document.getElementById("A").textContent || 0,
          B: document.getElementById("B").textContent || 0,
          C: document.getElementById("C").textContent || 0,
          U: document.getElementById("U").textContent || 0,
          AAA: document.getElementById("AAA").textContent || 0,
          totalStudents: document.getElementById("all").textContent || 0,
          passedStudents: document.getElementById("pass").textContent || 0,
          percentage: document.getElementById("studentsper").textContent || 0,
          universityAttainment:
            document.getElementById("university").textContent || 0,
        };

        const overallData = {
          cia70: document.getElementById("cia70").textContent || 0,
          ciaper: document.getElementById("ciaper").textContent || 0,
          ciauni: document.getElementById("ciauni").textContent || 0,
        };

        const overallAttainment = Array.from({ length: 5 }, (_, i) => ({
          co: `CO${i + 1}`,
          surveyIDA: document.getElementById(`ida${i + 1}`).textContent || 0,
          point1: document.getElementById(`point${i + 1}`).textContent || 0,
          point21: document.getElementById(`point2${i + 1}`).textContent || 0,
          ia: document.getElementById(`IA${i + 1}`).textContent || 0,
          totalPoint:
            document.getElementById(`totalPoint${i + 1}`).textContent || 0,
          max: document.getElementById(`max${i + 1}`).textContent || 0,
          allPoint: document.getElementById("allpoint").textContent || 0,
          average: document.getElementById("avarage").textContent || 0,
          allTotal: document.getElementById("alltotal").textContent || 0,
          ultraMax: document.getElementById("ultramax").textContent || 0,
        }));

        //console.log(JSON.stringify({ headerData, studentData, coSurveyData, eseData, overallData, overallAttainment,overallSummary }, null, 2));

        return {
          headerData,
          eseData,
          students,
          coSurveyData,
          overallData,
          overallAttainment,
        };
      }

function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== '') {
        const cookies = document.cookie.split(';');
        for (let i = 0; i < cookies.length; i++) {
            const cookie = cookies[i].trim();
            if (cookie.substring(0, name.length + 1) === (name + '=')) {
                cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                break;
            }
        }
    }
    return cookieValue;
}
const csrftoken = getCookie('csrftoken');
      async function saveToDatabase() {
        const data = collectData();
        console.log("Data to be sent to backend:", JSON.stringify(data));
        try {
          const response = await fetch("/department/theory/", {
            method: "POST",
            headers: { "Content-Type": "application/json","X-CSRFToken": csrftoken },
            body: JSON.stringify(data),
          });

          if (response.ok) {
            const result = await response.json();
            console.log("Server Response:", result);

            if (result.status === "success") {
              alert("Data saved successfully!");
                window.location.href="/department/";
            } else {
              alert(`Failed to save data: ${result.message}`);
            }
          } else {
            alert(`Server error: ${response.status}`);
          }
        } catch (error) {
          console.error("Unexpected error:", error);
          alert("An unexpected error occurred. Please try again.");
        }
      }


function handleFile() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: "A", raw: false });

        // Fill programme details
        document.getElementById('programme').value = sheet['A2'] ? sheet['A2'].v.replace('Programme: ', '') : ''; // Adjust cell reference
        document.getElementById('courseName').value = sheet['A3'] ? sheet['A3'].v.replace('Course Name: ', '') : '';
        document.getElementById('courseCode').value = sheet['A4'] ? sheet['A4'].v.replace('Course Code: ', '') : '';
        document.getElementById('academicYear').value = sheet['F2'] ? sheet['F3'].v : '';
        document.getElementById('semester').value = sheet['F3'] ? sheet['F3'].v : '';

        // Get number of students from cell I17 (adjust this as per your Excel format)
        const numStudents = parseInt(sheet['I17'] ? sheet['I17'].v : 0);
        document.getElementById('numStudents').value = numStudents;
        generateTable();

        // Fill student data (starting from row 7)
        for (let i = 0; i < numStudents; i++) {
            const regNo = sheet[`B${7 + i}`] ? sheet[`B${7 + i}`].v : '';
            const name = sheet[`C${7 + i}`] ? sheet[`C${7 + i}`].v : '';
            const ese = sheet[`D${7 + i}`] ? sheet[`D${7 + i}`].v : '';
            const cia = sheet[`E${7 + i}`] ? sheet[`E${7 + i}`].v : '';

            document.getElementById(`reg-${i + 1}`).value = regNo;
            document.getElementById(`name-${i + 1}`).value = name;
            document.getElementById(`ese-${i + 1}`).value = ese;
            document.getElementById(`cia-${i + 1}`).value = cia;

            // Manually trigger calculation functions
            calculateTotal1(i + 1);
            grade(i + 1);
            over();
            CIA(i+1);
        }

        // Fill Course Outcome Survey (COs) table (adjust cell references)
        for (let i = 1; i <= 5; i++) {
            document.getElementById(`Excellent${i}`).value = sheet[`I${i + 30}`] ? sheet[`I${i + 30}`].v : 0;
            document.getElementById(`Good${i}`).value = sheet[`J${i + 30}`] ? sheet[`J${i + 30}`].v : 0;
            document.getElementById(`Fair${i}`).value = sheet[`K${i + 30}`] ? sheet[`K${i + 30}`].v : 0;
            COs();
            over();
            updateStatistics();
        }
    };
    reader.readAsArrayBuffer(file);
}