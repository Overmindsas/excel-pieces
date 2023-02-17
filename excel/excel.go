package excel

import (
	"fmt"
	"strconv"
	"strings"
)

const (
	visocos  = 366
	noViscos = 365
)

type Excel struct{}

func ExcelNumberDate(source string) string {
	e := Excel{}
	if strings.TrimSpace(source) == "" {
		return ""
	}
	excelNumberDate, _ := strconv.Atoi(source)
	var (
		visocosBool, noVisocosBool bool

		startYear            = 1904
		excelNumberdateStart = 1462

		visocosYear  = 366
		noViscosYear = 365
	)

	sliceVisocos := []int{31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
	sliceNoViscos := []int{31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

	afterStartYear := 0
	for excelNumberdateStart <= excelNumberDate {
		if afterStartYear == 0 || afterStartYear%4 == 0 {
			excelNumberdateStart += visocos
		} else {
			excelNumberdateStart += noViscos
		}
		afterStartYear++
		startYear++
	}

	afterStartYear -= 1

	if afterStartYear%4 == 0 {
		visocosBool = true
	} else {
		noVisocosBool = true
	}

	startYear -= 1
	switch true {
	case noVisocosBool:
		return e.buildDate(noViscosYear, excelNumberdateStart, excelNumberDate, sliceNoViscos, startYear)
	case visocosBool:
		return e.buildDate(visocosYear, excelNumberdateStart, excelNumberDate, sliceVisocos, startYear)
	}
	return ""
}

func (e Excel) buildDate(visocosYear, excelNumberdateStart, excelNumberDate int, slice []int, startYear int) string {
	preRes := excelNumberdateStart - excelNumberDate - 1
	preRes = visocosYear - preRes
	for i := 0; i <= 12; i++ {
		preRes = preRes - slice[i]
		if preRes > slice[i+1] {
			continue
		}
		return e.validDate(preRes, i+2, startYear)
	}
	return ""
}

func (e Excel) validDate(day, month, year int) string {
	strDay, strMonth := "", ""
	if day < 10 {
		strDay = fmt.Sprintf("0%v", day)
	} else {
		strDay = fmt.Sprint(day)
	}
	if month < 10 {
		strMonth = fmt.Sprintf("0%v", month)
	} else {
		strMonth = fmt.Sprint(month)
	}
	return fmt.Sprintf("%v.%v.%v", strDay, strMonth, year)
}
