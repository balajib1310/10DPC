import streamlit as st


# -----------------------------
# Grade calculation
# -----------------------------
def get_grade(mark):
    if mark >= 90:
        return "A"
    elif mark >= 80:
        return "B"
    elif mark >= 70:
        return "C"
    elif mark >= 60:
        return "D"
    else:
        return "F"


# -----------------------------
# Page title
# -----------------------------
st.title("Student Grade Manager")


# -----------------------------
# Initialize session state
# -----------------------------
if "students" not in st.session_state:
    st.session_state.students = []


# -----------------------------
# Add Student Form
# -----------------------------
with st.form("add_student"):

    name = st.text_input("Name")

    mark = st.number_input(
        "Mark",
        min_value=0,
        max_value=100,
        step=1
    )

    submitted = st.form_submit_button("Add")

    if submitted:

        if name.strip() == "":
            st.error("Please enter the student's name.")

        else:
            grade = get_grade(mark)

            student = {
                "Name": name,
                "Mark": mark,
                "Grade": grade
            }

            st.session_state.students.append(student)

            st.success(f"{name} added successfully!")


# -----------------------------
# Display students and metrics
# -----------------------------
if st.session_state.students:

    st.subheader("Student Results")

    # Display table
    st.table(st.session_state.students)

    # Get all marks
    marks = [
        student["Mark"]
        for student in st.session_state.students
    ]

    # Calculate class statistics
    average = sum(marks) / len(marks)
    highest = max(marks)
    lowest = min(marks)

    # Display metrics
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "Average",
            f"{average:.1f}"
        )

    with col2:
        st.metric(
            "Highest",
            highest
        )

    with col3:
        st.metric(
            "Lowest",
            lowest
        )

else:

    st.info("No students added yet. Use the form above to add a student.")