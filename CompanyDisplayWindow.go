package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/widget"
)

type DisplayWindow struct {
	DataDisplay *widget.List
}

func BuildUI(window *DisplayWindow, mainWindow fyne.Window) {
	getCompanyName := widget.NewButton("Click here to get Campany name....", companyList)
    window.DataDisplay = widget.NewList(CreateListItem, UpdateListItem)
	topPane := container.NewVBox(window., getDataButton)
	leftPane := container.NewVSplit(topPane, window.DataDisplay)
}
