package main

import (
	"log"
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/widget"

	"github.com/xuri/excelize/v2"
)

// Declare global variables
var (
	file             *excelize.File  //Excel file
	firstColumn      []string        // Company names
	lastColumn       [][]string      // Remaining details of company
	companyList      *widget.List    // List widget for first column data
	lastColumnLabels []*widget.Entry // Entry widgets for last column data
	selectedIndex    int             // Variable to store the selected index
	mainWindow       fyne.Window     // Main window
)

func main() {
	var err error
	// Open Excel file
	file, err = excelize.OpenFile("Project2Data.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Get rows from a sheet in Excel file
	rows, err := file.GetRows("Comp490 Jobs")
	if err != nil {
		log.Fatal(err)
	}

	// Initialize Fyne app and main window
	displayApp := app.New()
	mainWindow = displayApp.NewWindow("Manpower Requirement")

	// Populate firstColumn and lastColumn with data from Excel
	if len(rows) > 0 {
		for _, row := range rows {
			if len(row) > 1 {
				firstColumn = append(firstColumn, row[0])
				lastColumn = append(lastColumn, row[1:])
			} else {
				log.Println("Row is empty or does not have enough elements")
			}
		}
	} else {
		log.Println("No rows found")
	}

	// Create list widget for firstColumn data
	companyList = widget.NewList(
		func() int { return len(firstColumn) },
		func() fyne.CanvasObject { return widget.NewLabel("") },
		func(id widget.ListItemID, item fyne.CanvasObject) {
			if label, ok := item.(*widget.Label); ok {
				label.SetText(firstColumn[id])
			}
		},
	)

	// Create entry widgets for lastColumn data
	lastColumnLabels = make([]*widget.Entry, len(lastColumn[0]))
	for i := range lastColumnLabels {
		lastColumnLabels[i] = widget.NewEntry()
	}

	// Create a container for lastColumn data entry widgets
	lastColumnContainer := container.NewVBox()
	for _, entry := range lastColumnLabels {
		lastColumnContainer.Add(entry)
	}

	// Update lastColumn data when firstColumn item is clicked
	companyList.OnSelected = func(id int) {
		if id >= 0 && id < len(lastColumn) {
			for i, val := range lastColumn[id] {
				lastColumnLabels[i].SetText(val)
			}
			selectedIndex = id // Update the selected index
		}
	}

	// Create buttons for CRUD operations
	buttons := createButtonWidgets()
	// Create content layout
	content := container.NewHSplit(companyList, container.NewVBox(lastColumnContainer))
	// Create bottom layout for buttons
	bottom := container.NewHBox(buttons)

	// Set content and resize main window
	mainWindow.SetContent(container.NewBorder(content, bottom, nil, nil))
	mainWindow.Resize(fyne.NewSize(800, 500))
	mainWindow.ShowAndRun()
}

// Function to create buttons for CRUD operations
func createButtonWidgets() fyne.CanvasObject {

	// Create "Create" button
	addButton := widget.NewButton("Create", func() {
		addData()
	})

	// Create "Delete" button
	deleteButton := widget.NewButton("Delete", func() {
		deleteData(selectedIndex)
	})

	// Create "Update" button
	updateButton := widget.NewButton("Update", func() {
		updateData()
	})

	// Return buttons layout
	return container.NewHBox(addButton, deleteButton, updateButton)
}

// Function to add data
func addData() {
	var rowData []string
	for _, entry := range lastColumnLabels {
		rowData = append(rowData, entry.Text)
	}
	lastColumn = append(lastColumn, rowData)
	companyList.Refresh()
	appendDataToExcel(rowData)
	clearEntryWidgets()
}

// Function to delete data
func deleteData(id int) {
	if id >= 0 && id < len(lastColumn) {
		lastColumn = append(lastColumn[:id], lastColumn[id+1:]...)
		companyList.Refresh()
		deleteRowFromExcel(id + 2)
		clearEntryWidgets()
	}
}

// Function to update data
func updateData() {
	id := selectedIndex
	if id >= 0 && id < len(lastColumn) {
		var rowData []string
		for _, entry := range lastColumnLabels {
			rowData = append(rowData, entry.Text)
		}
		lastColumn[id] = rowData
		companyList.Refresh()
		updateRowInExcel(id+2, rowData)
		clearEntryWidgets()
	}
}

// Function to delete row from Excel
func deleteRowFromExcel(row int) {
	file.RemoveRow("Comp490 Jobs", row)
	saveExcelFile()
}

// Function to update row from Excel
func updateRowInExcel(row int, data []string) {
	for i, value := range data {
		cell := ToAlphaString(i) + strconv.Itoa(row)
		file.SetCellValue("Comp490 Jobs", cell, value)
	}
	saveExcelFile()
}

// Function to save Excel file
func saveExcelFile() {
	if err := file.Save(); err != nil {
		log.Fatal(err)
	}
}

// Function to clear entry widgets
func clearEntryWidgets() {
	for _, entry := range lastColumnLabels {
		entry.SetText("")
	}
}

// Function to convert integer to alphabet string
func ToAlphaString(i int) string {
	if i < 0 {
		return ""
	}

	var result string
	for i >= 0 {
		result = string('A'+i%26) + result
		i = i/26 - 1
	}
	return result
}

// Function to append data to Excel file
func appendDataToExcel(data []string) {
	row := len(lastColumn) + 1
	for i, value := range data {
		cell := ToAlphaString(i) + strconv.Itoa(row)
		file.SetCellValue("Comp490 Jobs", cell, value)
	}
	saveExcelFile()
}
