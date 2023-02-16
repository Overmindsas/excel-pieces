package excel

import (
	"fmt"
	"strconv"
	"strings"
)

func ExcelNumberDate(source string) string {
	if strings.TrimSpace(source) == "" {
		return ""
	}
	excelNumberDate, _ := strconv.Atoi(source)
	var (
		visocosBool   bool
		noVisocosBool bool

		startYear            = 2008
		excelNumberdateStart = 39448

		visocosYear  = 366
		noViscosYear = 365
	)

	const (
		visocos  = 366
		noViscos = 365
	)

	sliceVisocos := []int{31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
	sliceNoViscos := []int{31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

	counter := 0
	for excelNumberdateStart <= excelNumberDate {
		if counter == 0 || counter%4 == 0 {
			excelNumberdateStart += visocos
		} else {
			excelNumberdateStart += noViscos
		}
		counter++
		startYear++
	}

	counter -= 1

	if counter%4 == 0 {
		visocosBool = true
	} else {
		noVisocosBool = true
	}

	startYear -= 1
	switch true {
	case noVisocosBool:
		preRes := excelNumberdateStart - excelNumberDate - 1
		preRes = noViscosYear - preRes
		for i := 0; i <= 12; i++ {
			preRes = preRes - sliceNoViscos[i]
			if preRes > sliceNoViscos[i+1] {
				continue
			}
			return validDate(preRes, i+2, startYear)
		}
	case visocosBool:
		preRes := excelNumberdateStart - excelNumberDate - 1
		preRes = visocosYear - preRes
		for i := 0; i < 12; i++ {
			preRes = preRes - sliceVisocos[i]
			if preRes > sliceVisocos[i+1] {
				continue
			}
			return validDate(preRes, i+2, startYear)
		}
	}
	return ""
}

func validDate(day, month, year int) string {
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
