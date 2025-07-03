# Gantt-Chart-Tool
The Python Gantt Chart Tool is a Tkinter desktop app for project management. Users can create, visualize, filter, and sort tasks by dates, priority, and status. It features AI assistance (Google Gemini) for task descriptions, sub-tasks, risks, and status updates. Data save/load (JSON/CSV), Excel export, undo/redo, theme customization included.

Python Gantt Chart Tool
A desktop application built with Tkinter for efficient project task management and visualization, featuring AI assistance powered by Google Gemini.

Table of Contents
Overview

Features

Setup and Installation

How to Run the Application

Core Features & Usage

Contributing

License

Overview
This Python-based Gantt Chart Tool is designed to help individuals and small teams manage and visualize project tasks effectively. It provides a user-friendly graphical interface to add, edit, delete, and track tasks, along with powerful filtering, sorting, and AI-driven assistance for task content generation.

Features
Task Management: Create and manage tasks with comprehensive details including name, epic number, start/end dates, custom color, priority (Low, Medium, High, Critical), status (Not Started, In Progress, Completed, Blocked), and milestone designation.

Dynamic Gantt Chart Visualization: Visually represent tasks on a dynamic Gantt chart, clearly showing their duration, overlaps, and current week highlight.

Filtering & Sorting: Easily filter tasks by Epic number or date range, and sort them by Task Name, Start Date, End Date, Priority, or Status for better focus.

AI Assistance (Google Gemini): Leverage integrated AI capabilities to:

Expand on task descriptions.

Generate actionable sub-tasks.

Brainstorm potential risks and mitigation strategies.

Draft concise status updates.

Data Persistence: Automatically saves and loads project data to/from a gantt_tasks.json file. Supports manual saving to JSON and loading from JSON/CSV.

Export Capabilities: Export your task data, including a visual representation of the Gantt chart, to an Excel (.xlsx) spreadsheet.

Undo/Redo: Conveniently revert or re-apply changes to your task list.

Customizable Theme: Choose from various Tkinter themes to personalize the application's appearance.

Setup and Installation
To get the application up and running, follow these steps:

Install Python:
Ensure you have Python installed on your system. Python 3.8 or newer is recommended. Download it from python.org.

Install Required Libraries:
Open your terminal or command prompt and run the following command to install all necessary Python packages:

pip install tkcalendar openpyxl google-generativeai langchain-google-genai

Obtain a Google Gemini API Key:
The AI Assist feature requires a Google Gemini API key.

Go to the Google AI Studio or Google Cloud Console to obtain your API key.

For security, it is crucial to set this as an environment variable rather than hardcoding it into the script, especially when sharing your code publicly.

On Windows (Command Prompt):
To set persistently (requires restarting terminal):

setx GOOGLE_API_KEY "YOUR_ACTUAL_API_KEY"

For the current session only:

set GOOGLE_API_KEY=YOUR_ACTUAL_API_KEY

On macOS/Linux (Terminal):
For the current session only:

export GOOGLE_API_KEY="YOUR_ACTUAL_API_KEY"

To set persistently (add this line to your ~/.bashrc, ~/.zshrc, or ~/.profile file, then source the file or restart your terminal):

export GOOGLE_API_KEY="YOUR_ACTUAL_API_KEY"

Remember to replace "YOUR_ACTUAL_API_KEY" with your actual API key.

How to Run the Application
Save the Code: Save the main application code (e.g., gantt_chart_app.py) to a file on your computer.

Open Terminal/Command Prompt: Navigate to the directory where you saved the .py file.

Execute the Script:

python gantt_chart_app.py

The Gantt Chart Tool's graphical user interface will launch.

Core Features & Usage
Adding/Updating Tasks: Fill in the "Task Details" fields (Name, Epic #, Start/End Dates, Color, Priority, Status, Is Milestone) and click "Add Task". To update, select a task on the chart (using the pencil icon), modify details, and click "Update Task".

AI Assist: Click the "AI Assist" button next to the Task Name. In the new window, enter/confirm the task name, select an "AI Action" (e.g., "Expand Description", "Generate Sub-tasks"), click "Generate", and then "Apply to Main Task" to insert the AI-generated content.

Editing/Deleting: Use the "✏️" (edit) and "🗑️" (delete) buttons next to each task on the Gantt chart.

Filtering: Use the "Filter" section to narrow down tasks by "Epic #" or a "Start Date" / "End Date" range. Click "Clear Filters" to reset.

Sorting: Click the buttons in the "Sort" section (e.m., "Task Name", "Start Date") to reorder tasks on the chart.

Undo/Redo: Use the "Undo" and "Redo" buttons to manage your task list history.

Saving/Loading Data: The app auto-saves to gantt_tasks.json. Use "Save Data (JSON)" for manual saves or "Load Data (JSON/CSV)" to import from files.

Export to Excel: Click "Export to Excel" to generate an .xlsx file containing your task data and a basic chart representation.

Changing Theme: Select a theme from the "Theme" dropdown at the bottom right of the window.

Contributing
Contributions are welcome! If you have suggestions for improvements, new features, or bug fixes, feel free to:

Fork the repository.

Create a new branch (git checkout -b feature/YourFeatureName).

Make your changes.

Commit your changes (git commit -m 'Add new feature').

Push to the branch (git push origin feature/YourFeatureName).

Open a Pull Request.

License
This project is open-source and available under the MIT License.
