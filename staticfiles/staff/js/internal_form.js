function generateTable() {
    const numStudents = parseInt(document.getElementById('numStudents').value);
    const workingDays = parseInt(document.getElementById('totalWorkingDays').value) || 0; // Get initial value
    const tbody = document.getElementById('studentsTable').querySelector('tbody');
    tbody.innerHTML = '';

    for (let i = 1; i <= numStudents; i++) {
        const row = `
            <tr>
                <td><input type="number" class="regno" id="reg-${i}"></td>
                <td><input type="text" class="student-name" id="name-${i}"></td>
                <td><input type="number" class="numeric-field" id="present-${i}" oninput="calculateAttendance(${i})"></td>
                <td><input type="number" class="numeric-field" id="workingdays-${i}" value="${workingDays}" readonly></td>
                <td><input type="number" class="numeric-field" id="percentage-${i}" readonly></td>
                <td><input type="number" class="numeric-field" id="mark-${i}" oninput="calculateOverallTotal(${i})" readonly></td>
                <td><input type="number" class="numeric-field" id="assignment-${i}" oninput="calculateOverallTotal(${i})"></td>
                <td><input type="number" class="numeric-field" id="seminar-${i}" oninput="calculateOverallTotal(${i})"></td>
                <td><input type="number" class="numeric-field" id="1internal-${i}" oninput="calculateTotalMarks(${i})"></td>
                <td><input type="number" class="numeric-field" id="2internal-${i}" oninput="calculateTotalMarks(${i})"></td>
                <td><input type="number" class="numeric-field" id="model-${i}" oninput="calculateTotalMarks(${i})"></td>
                <td><input type="number" class="numeric-field" id="total-${i}"oninput="calculateOverallTotal(${i})" readonly></td>
                <td><input type="number" class="numeric-field" id="overall-${i}" readonly></td>
            </tr>`;
        tbody.innerHTML += row;
    }
    enableArrowKeyNavigation();
}

function updateWorkingDays() {
    const totalWorkingDays = parseInt(document.getElementById('totalWorkingDays').value) || 0;
    const numStudents = parseInt(document.getElementById('numStudents').value);

    for (let i = 1; i <= numStudents; i++) {
        // Update working days value in each row
        document.getElementById(`workingdays-${i}`).value = totalWorkingDays;

        // Recalculate attendance and marks for each student
        calculateAttendance(i);
    }
}


function calculateAttendance(index) {
    const presentDays = parseFloat(document.getElementById(`present-${index}`).value) || 0;
    const workingDays = parseFloat(document.getElementById(`workingdays-${index}`).value) || 1; // Prevent division by zero
    const percentage = Math.round((presentDays / workingDays) * 100); // Rounded to the nearest whole number

    // Update the percentage field
    document.getElementById(`percentage-${index}`).value = percentage;

    // Calculate marks based on percentage
    let mark = 0;
    if (percentage >= 90 && percentage <= 100) {
        mark = 5;
    } else if (percentage >= 80 && percentage <= 89) {
        mark = 4;
    } else if (percentage >= 70 && percentage <= 79) {
        mark = 3;
    } else if (percentage >= 60 && percentage <= 69) {
        mark = 2;
    }

    // Update the mark field
    document.getElementById(`mark-${index}`).value = mark;
}


        function enableArrowKeyNavigation() {
            const inputs = document.querySelectorAll('input');
            inputs.forEach((input, index) => {
                input.addEventListener('keydown', (e) => {
                    let nextInput;
                    const totalInputs = inputs.length;
                    const cols = 13;

                    switch (e.key) {
                        case 'ArrowRight':
                            nextInput = inputs[(index + 1) % totalInputs];
                            break;
                        case 'ArrowLeft':
                            nextInput = inputs[(index - 1 + totalInputs) % totalInputs];
                            break;
                        case 'ArrowUp':
                            nextInput = inputs[(index - cols + totalInputs) % totalInputs];
                            break;
                        case 'ArrowDown':
                            nextInput = inputs[(index + cols) % totalInputs];
                            break;
                        default:
                            return;
                    }
                    nextInput.focus();
                    e.preventDefault(); // Prevent default scroll behavior
                });
            });
        }

        function calculateTotalMarks(index) {
    const internal1 = parseFloat(document.getElementById(`1internal-${index}`).value) || 0;
    const internal2 = parseFloat(document.getElementById(`2internal-${index}`).value) || 0;
    const modelExam = parseFloat(document.getElementById(`model-${index}`).value) || 0;

    // Calculate and round the total score to the nearest whole number
    const totalScore = Math.round(((internal1 + internal2 + modelExam) / 175) * 10);

    // Update the total field with the rounded value
    document.getElementById(`total-${index}`).value = totalScore;

    // Trigger overall total update automatically
    calculateOverallTotal(index);
}



function calculateTotalMarks(index) {
    const internal1 = parseFloat(document.getElementById(`1internal-${index}`).value) || 0;
    const internal2 = parseFloat(document.getElementById(`2internal-${index}`).value) || 0;
    const modelExam = parseFloat(document.getElementById(`model-${index}`).value) || 0;

    // Calculate and round to the nearest whole number
    const totalScore = ((internal1 + internal2 + modelExam) / 175) * 10;
    const roundedTotal = Math.round(totalScore); // This gives a whole number

    document.getElementById(`total-${index}`).value = roundedTotal;

    // Trigger overall total update automatically
    calculateOverallTotal(index);
}


function calculateOverallTotal(index) {
    const marks = parseFloat(document.getElementById(`mark-${index}`).value) || 0;
    const assignment = parseFloat(document.getElementById(`assignment-${index}`).value) || 0;
    const seminar = parseFloat(document.getElementById(`seminar-${index}`).value) || 0;
    const total = parseFloat(document.getElementById(`total-${index}`).value) || 0;

    // Sum of all marks
    const overallTotal = marks + assignment + seminar + total;

    // Update the overall total field
    document.getElementById(`overall-${index}`).value = overallTotal;
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

async function submitData() {
    // Fetch general details
    const headerData = {
        programme: document.getElementById("programme").value,
        courseName: document.getElementById("courseName").value,
        courseCode: document.getElementById("courseCode").value,
        academicYear: document.getElementById("academicYear").value,
        semester: document.getElementById("semester").value,
        totalWorkingDays: document.getElementById("totalWorkingDays").value,
    };

    const students = [];
    const rows = document.getElementById("studentsTable").querySelectorAll("tbody tr");

    rows.forEach((row, index) => {
        students.push({
            reg_no: document.getElementById(`reg-${index + 1}`).value,
            student_name: document.getElementById(`name-${index + 1}`).value,
            days_present: parseInt(document.getElementById(`present-${index + 1}`).value) || 0,
            working_days: parseInt(document.getElementById(`workingdays-${index + 1}`).value) || 0,
            percentage: parseFloat(document.getElementById(`percentage-${index + 1}`).value) || 0,
            marks: parseFloat(document.getElementById(`mark-${index + 1}`).value) || 0,
            assignment: parseFloat(document.getElementById(`assignment-${index + 1}`).value) || 0,
            seminar: parseFloat(document.getElementById(`seminar-${index + 1}`).value) || 0,
            internal1: parseFloat(document.getElementById(`1internal-${index + 1}`).value) || 0,
            internal2: parseFloat(document.getElementById(`2internal-${index + 1}`).value) || 0,
            model_exam: parseFloat(document.getElementById(`model-${index + 1}`).value) || 0,
            total: parseFloat(document.getElementById(`total-${index + 1}`).value) || 0,
            overall_total: parseFloat(document.getElementById(`overall-${index + 1}`).value) || 0,
        });
    });

    const data = { headerData, students };

    try {
        const response = await fetch("/department/internal/", {
            method: "POST",
            headers: { "Content-Type": "application/json" , "X-CSRFToken": csrftoken},
            body: JSON.stringify(data),
        });
        if (response.ok) {
            const result = await response.json();
            console.log("Server Response:", result); // Log the response

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
    const fileInput = document.getElementById('uploadExcel');
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function (event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                alert("No sheets found in the Excel file.");
                return;
            }

            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!sheet) {
                alert("Sheet data not found.");
                return;
            }

            // ✅ Safely fill metadata
            document.getElementById('programme').value =
                sheet['A3']?.v?.toString().replace('Programme: ', '').trim() || '';

            document.getElementById('courseName').value =
                sheet['H3']?.v?.toString().replace('Course Name: ', '').trim() || '';

            document.getElementById('courseCode').value =
                sheet['M3']?.v?.toString().replace('Course Code: ', '').trim() || '';

            if (sheet['D3']?.v) {
                let yearSem = sheet['D3'].v.replace('Year & Sem:', '').trim();
                let parts = yearSem.split('&').map(p => p.trim());
                document.getElementById('academicYear').value = parts[0] || '';
                document.getElementById('semester').value = parts[1] || '';
            } else {
                document.getElementById('academicYear').value = '';
                document.getElementById('semester').value = '';
            }

            document.getElementById('totalWorkingDays').value = sheet['E6']?.v || '';

            // ✅ Count students (row 6 onwards)
            let lastRow = 6;
            while (sheet[`C${lastRow}`] && sheet[`C${lastRow}`].v) {
                lastRow++;
            }
            const numStudents = lastRow - 6;
            document.getElementById('numStudents').value = numStudents > 0 ? numStudents : 0;

            // ✅ Ensure table exists before filling
            if (typeof generateTable === "function") {
                generateTable();
            }

            // ✅ Fill student data
            for (let i = 0; i < numStudents; i++) {
                let row = 6 + i;
                let studentIndex = i + 1;

                const fields = [
                    { id: `reg-${studentIndex}`, value: sheet[`B${row}`]?.v || '' },
                    { id: `name-${studentIndex}`, value: sheet[`C${row}`]?.v || '' },
                    { id: `present-${studentIndex}`, value: sheet[`D${row}`]?.v || '' },
                    { id: `assignment-${studentIndex}`, value: sheet[`H${row}`]?.v || '' },
                    { id: `seminar-${studentIndex}`, value: sheet[`I${row}`]?.v || '' },
                    { id: `1internal-${studentIndex}`, value: sheet[`J${row}`]?.v || '' },
                    { id: `2internal-${studentIndex}`, value: sheet[`K${row}`]?.v || '' },
                    { id: `model-${studentIndex}`, value: sheet[`L${row}`]?.v || '' }
                ];

                fields.forEach(field => {
                    let inputElement = document.getElementById(field.id);
                    if (inputElement) {
                        inputElement.value = field.value;
                        inputElement.dispatchEvent(new Event('input'));
                    }
                });

                // ✅ Trigger calculations only if defined
                if (typeof calculateAttendance === "function") calculateAttendance(studentIndex);
                if (typeof calculateTotalMarks === "function") calculateTotalMarks(studentIndex);
                if (typeof calculateOverallTotal === "function") calculateOverallTotal(studentIndex);
            }

        } catch (error) {
            alert("Error reading Excel file: " + error.message);
            console.error(error);
        }
    };

    reader.readAsArrayBuffer(file);
}

