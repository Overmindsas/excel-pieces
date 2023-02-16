package excel

import (
	"testing"

	"github.com/test-go/testify/assert"
)

func TestExcelNumberdate(t *testing.T) {
	assertion := assert.New(t)
	t.Log("Test ExcelNumberdate")
	{
		testID := 0
		t.Logf("Test №%v:", testID)
		{
			date := ExcelNumberDate("")
			assertion.Equal(date, "")
		}

		testID++
		t.Logf("Test №%v:", testID)
		{
			date := ExcelNumberDate("44668")
			assertion.Equal(date, "17.04.2022")
		}

		testID++
		t.Logf("Test №%v:", testID)
		{
			date := ExcelNumberDate("41041")
			assertion.Equal(date, "12.05.2012")
		}

		testID++
		t.Logf("Test №%v:", testID)
		{
			date := ExcelNumberDate("43159")
			assertion.Equal(date, "28.02.2018")
		}
	}
}
