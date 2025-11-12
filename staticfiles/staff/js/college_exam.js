function getTotalMarks() {
    const examName = document.getElementById("exam_name").value;
    if (examName === "I Internal" || examName === "II Internal") {
        return 50;
    } else if (examName === "model exam") {
        return 75;
    }
    return 0; // Default case
}

function generateTable() {
    const numStudents = parseInt(document.getElementById('numStudents').value) || 0;
    const totalMarks = getTotalMarks();
    const tbody = document.getElementById('studentsTable').querySelector('tbody');
    tbody.innerHTML = '';

    for (let i = 1; i <= numStudents; i++) {
        const row = `
            <tr>
                <td><input type="number" class="regno" id="reg-${i}" placeholder="Reg. No"></td>
                <td><input type="text" class="student-name" id="name-${i}" placeholder="Student Name"></td>
                <td><input type="number" class="numeric-field" id="marks-${i}" oninput="calculatePercentage(${i})"></td>
                <td><input type="number" class="numeric-field total-marks" id="total-${i}" value="${totalMarks}" readonly></td>
                <td><input type="text" class="numeric-field" id="percentage-${i}" readonly></td>
            </tr>`;
        tbody.innerHTML += row;
    }
    setTimeout(() => {
    const inputs = tbody.querySelectorAll('input');
    inputs.forEach((input, index) => {
        input.addEventListener('keydown', (e) => {
            if (['ArrowRight', 'ArrowLeft', 'ArrowDown', 'ArrowUp'].includes(e.key)) {
                e.preventDefault();
                let rowLength = 5;
                let newIndex = index;
                switch (e.key) {
                    case 'ArrowRight': newIndex = index + 1; break;
                    case 'ArrowLeft': newIndex = index - 1; break;
                    case 'ArrowDown': newIndex = index + rowLength; break;
                    case 'ArrowUp': newIndex = index - rowLength; break;
                }
                if (inputs[newIndex]) {
                    inputs[newIndex].focus();
                }
            }
        });
    });
}, 0);

}

function calculatePercentage(index) {
    const marks = parseFloat(document.getElementById(`marks-${index}`).value) || 0;
    const total = parseFloat(document.getElementById(`total-${index}`).value) || 0;

    const percentage = total > 0 ? ((marks / total) * 100).toFixed(2) : "0.00";
    document.getElementById(`percentage-${index}`).value = percentage;
}

function updateTotalMarks() {
    const totalMarks = getTotalMarks();
    const totalCells = document.querySelectorAll(".total-marks");

    totalCells.forEach((cell, index) => {
        cell.value = totalMarks; // Update total marks column
        calculatePercentage(index + 1); // Recalculate percentage
    });
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
    const headerData = {
        programme: document.getElementById('programme').value,
        courseName: document.getElementById('courseName').value,
        courseCode: document.getElementById('courseCode').value,
        academicYear: document.getElementById('academicYear').value,
        examName: document.getElementById('exam_name').value, // fixed ID
        semester: document.getElementById('semester').value,
        numStudents: parseInt(document.getElementById('numStudents').value) || 0,
    };

    const studentsData = [];
    let numStudents = parseInt(document.getElementById('numStudents').value) || 0;

    for (let i = 1; i <= numStudents; i++) {
        studentsData.push({
            regNo: document.getElementById(`reg-${i}`).value,
            studentName: document.getElementById(`name-${i}`).value,
            marks: parseInt(document.getElementById(`marks-${i}`).value) || 0,
            total: parseInt(document.getElementById(`total-${i}`).value) || 0,
            percentage: parseFloat(document.getElementById(`percentage-${i}`).value) || 0,
        });
    }

    const data = { headerData, studentsData };

    try {
        const response = await fetch("/department/clg_exam/", {
            method: "POST",
            headers: { "Content-Type": "application/json", "X-CSRFToken": csrftoken },
            body: JSON.stringify(data),
        });

        if (response.ok) {
            const result = await response.json();
            console.log("Server Response:", result);

            if (result.status === "success") {
                alert("Data saved successfully!");
                window.location.href = "/department/";
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

