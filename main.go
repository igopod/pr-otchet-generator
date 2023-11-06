package main

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/net/html/charset"
)

type otchet struct {
	ФайлОСЗН       xml.Name `xml:"ФайлОСЗН"`
	ИмяФайла       string   `xml:"ИмяФайла"`
	ЗаголовокФайла struct {
		ВерсияФормата             string `xml:"ВерсияФормата"`
		ТипФайла                  string `xml:"ТипФайла"`
		ПрограммаПодготовкиДанных struct {
			НазваниеПрограммы string `xml:"НазваниеПрограммы"`
			Версия            string `xml:"Версия"`
		} `xml:"ПрограммаПодготовкиДанных"`
		ИсточникДанных string `xml:"ИсточникДанных"`
	} `xml:"ЗаголовокФайла"`
	ПачкаИсходящихДокументов struct {
		ДоставочнаяОрганизация string `xml:"ДоставочнаяОрганизация,arrt"`

		ИСХОДЯЩАЯ_ОПИСЬ struct {
			СоставительПачки struct {
				НалоговыйНомер struct {
					ИНН string `xml:"ИНН"`
					КПП string `xml:"КПП"`
				} `xml:"НалоговыйНомер"`

				НаименованиеОрганизации string `xml:"НаименованиеОрганизации"`

				РегистрационныйНомер string `xml:"РегистрационныйНомер"`
			} `xml:"СоставительПачки"`

			СоставДокументов struct {
				Количество int `xml:"Количество"`

				НаличиеДокументов struct {
					ТипДокумента string `xml:"ТипДокумента"`
					Количество   int    `xml:"Количество"`
				} `xml:"НаличиеДокументов"`
			} `xml:"СоставДокументов"`

			ТерриториальныйОрган struct {
				НаименованиеОрганизации string `xml:"НаименованиеОрганизации"`
				РегистрационныйНомер    string `xml:"РегистрационныйНомер"`
			} `xml:"ТерриториальныйОрган"`

			ОрганизацияСформировавшаяДокумент struct {
				НаименованиеОрганизации string `xml:"НаименованиеОрганизации"`
				РегистрационныйНомер    string `xml:"РегистрационныйНомер"`
			} `xml:"ОрганизацияСформировавшаяДокумент"`

			ТипМассиваПоручений string `xml:"ТипМассиваПоручений"`
			Месяц               int    `xml:"Месяц"`
			Год                 int    `xml:"Год"`
			ДатаФормирования    string `xml:"ДатаФормирования"`
			Должность           string `xml:"Должность"`
			Руководитель        string `xml:"Руководитель"`
		} `xml:"ИСХОДЯЩАЯ_ОПИСЬ"`

		ПОДТВЕРЖДЕНИЕ_О_ПОЛУЧЕНИИ_МАССИВА_И_СОПРОВОДИТЕЛЬНОЙ_ОПИСИ struct {
			НомерОписиПоручений int `xml:"НомерОписиПоручений"`
			Количество          int `xml:"Количество"`

			СведенияОмассивеПоручений []struct {
				ТипСтроки                      string  `xml:"ТипСтроки"`
				НомерОПС                       string  `xml:"НомерОПС"`
				КоличествоПоручений            int     `xml:"КоличествоПоручений"`
				СуммаПоМассиву                 float64 `xml:"СуммаПоМассиву"`
				СистемныйНомерМассиваПоручений string  `xml:"СистемныйНомерМассиваПоручений"`
				ПодтверждениеВполучении        string  `xml:"ПодтверждениеВполучении"`
			} `xml:"СведенияОмассивеПоручений"`
			ДатаВыдачиДокумента string `xml:"ДатаВыдачиДокумента"`
		} `xml:"ПОДТВЕРЖДЕНИЕ_О_ПОЛУЧЕНИИ_МАССИВА_И_СОПРОВОДИТЕЛЬНОЙ_ОПИСИ"`
	} `xml:"ПачкаИсходящихДокументов"`
}

func main() {
	input := os.Args[1]

	outputDir := filepath.Dir(input)
	output := filepath.Join(outputDir, "Отчет.xlsx")

	xmlFile, err := os.Open(input)
	if err != nil {
		log.Printf("error: %v", err)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		log.Printf("error: %v", err)
	}

	data := new(otchet)

	reader := bytes.NewReader(byteValue)
	decoder := xml.NewDecoder(reader)
	decoder.CharsetReader = charset.NewReaderLabel
	err = decoder.Decode(&data)
	if err != nil {
		log.Printf("error: %v", err)
	}

	xlsxFile := excelize.NewFile()
	defer xlsxFile.Close()

	err = xlsxFile.SetDefaultFont("Liberation Serif")
	if err != nil {
		log.Println(err)
	}

	sheet := "Sheet1"

	headerStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
		Font: &excelize.Font{
			Italic: false,
			Size:   12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	headerBoldStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{},
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	tableHeaderStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Italic: true,
			Size:   12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	tableNumbersStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{
			Size: 12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	tableOPSStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Size: 12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	tableFinalStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			// Horizontal: "center",
			// Vertical:   "center",
		},
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
	})
	if err != nil {
		log.Printf("error: %v", err)
	}

	docType := data.ПачкаИсходящихДокументов.ИСХОДЯЩАЯ_ОПИСЬ.СоставДокументов.НаличиеДокументов.ТипДокумента
	docType = strings.ReplaceAll(docType, "_", " ")
	orgName := data.ПачкаИсходящихДокументов.ИСХОДЯЩАЯ_ОПИСЬ.СоставительПачки.НаименованиеОрганизации
	arrType := data.ПачкаИсходящихДокументов.ИСХОДЯЩАЯ_ОПИСЬ.ТипМассиваПоручений
	docDate := data.ПачкаИсходящихДокументов.ПОДТВЕРЖДЕНИЕ_О_ПОЛУЧЕНИИ_МАССИВА_И_СОПРОВОДИТЕЛЬНОЙ_ОПИСИ.ДатаВыдачиДокумента
	arrPoruchenii := data.ПачкаИсходящихДокументов.ПОДТВЕРЖДЕНИЕ_О_ПОЛУЧЕНИИ_МАССИВА_И_СОПРОВОДИТЕЛЬНОЙ_ОПИСИ.СведенияОмассивеПоручений

	var docMonthYear string
	m := strings.Split(docDate, ".")[1]
	y := strings.Split(docDate, ".")[2]

	switch m {
	case "01":
		docMonthYear = strings.Join([]string{"за январь", y, "года"}, " ")
	case "02":
		docMonthYear = strings.Join([]string{"за февраль", y, "года"}, " ")
	case "03":
		docMonthYear = strings.Join([]string{"за март", y, "года"}, " ")
	case "04":
		docMonthYear = strings.Join([]string{"за апрель", y, "года"}, " ")
	case "05":
		docMonthYear = strings.Join([]string{"за май", y, "года"}, " ")
	case "06":
		docMonthYear = strings.Join([]string{"за июнь", y, "года"}, " ")
	case "07":
		docMonthYear = strings.Join([]string{"за июль", y, "года"}, " ")
	case "08":
		docMonthYear = strings.Join([]string{"за август", y, "года"}, " ")
	case "09":
		docMonthYear = strings.Join([]string{"за сентябрь", y, "года"}, " ")
	case "10":
		docMonthYear = strings.Join([]string{"за октябрь", y, "года"}, " ")
	case "11":
		docMonthYear = strings.Join([]string{"за ноябрь", y, "года"}, " ")
	case "12":
		docMonthYear = strings.Join([]string{"за декабрь", y, "года"}, " ")
	default:
		docMonthYear = "Ошибка! Месяц не определен"
	}

	xlsxFile.MergeCell(sheet, "A3", "D3")
	xlsxFile.SetRowHeight(sheet, 3, 30)
	xlsxFile.MergeCell(sheet, "A4", "D4")
	xlsxFile.MergeCell(sheet, "A5", "D5")

	xlsxFile.SetColWidth(sheet, "A", "A", 11.86)
	xlsxFile.SetColWidth(sheet, "B", "B", 17.29)
	xlsxFile.SetColWidth(sheet, "C", "C", 23.57)
	xlsxFile.SetColWidth(sheet, "D", "D", 29.00)

	xlsxFile.SetCellValue(sheet, "A1", docDate)
	xlsxFile.SetCellValue(sheet, "A3", docType)
	xlsxFile.SetCellValue(sheet, "A4", orgName)
	xlsxFile.SetCellValue(sheet, "A5", docMonthYear)
	xlsxFile.SetCellValue(sheet, "A7", arrType)

	xlsxFile.SetCellValue(sheet, "A9", "№")
	xlsxFile.SetCellValue(sheet, "B9", "ОПС")
	xlsxFile.SetCellValue(sheet, "C9", "Количество поручений")
	xlsxFile.SetCellValue(sheet, "D9", "Cумма")

	currentRowIndex := 9
	tableStartRow := currentRowIndex
	for index, element := range arrPoruchenii {
		currentRowIndex += 1

		if index == 0 {
			xlsxFile.SetCellValue(sheet, "A8", element.СистемныйНомерМассиваПоручений)
		}

		xlsxFile.SetCellValue(sheet, fmt.Sprintf("A%d", currentRowIndex), index+1)

		if element.ТипСтроки == "ИТОГО" {
			xlsxFile.SetCellValue(sheet, fmt.Sprintf("A%d", currentRowIndex), element.ТипСтроки)
		}

		xlsxFile.SetCellValue(sheet, fmt.Sprintf("B%d", currentRowIndex), element.НомерОПС)
		xlsxFile.SetCellValue(sheet, fmt.Sprintf("C%d", currentRowIndex), element.КоличествоПоручений)
		xlsxFile.SetCellValue(sheet, fmt.Sprintf("D%d", currentRowIndex), element.СуммаПоМассиву)
	}
	tableEndRow := currentRowIndex

	xlsxFile.SetCellValue(sheet, fmt.Sprintf("A%d", currentRowIndex+2), "Руководитель ИВЦ")
	xlsxFile.SetCellValue(sheet, fmt.Sprintf("A%d", currentRowIndex+4), "Сформировано")

	xlsxFile.SetCellStyle(sheet, "A3", "A5", headerStyle)
	xlsxFile.SetCellStyle(sheet, "A7", "A8", headerBoldStyle)
	xlsxFile.SetCellStyle(sheet, fmt.Sprintf("%s%d", "A", tableStartRow), fmt.Sprintf("%s%d", "D", tableEndRow), tableNumbersStyle)
	xlsxFile.SetCellStyle(sheet, fmt.Sprintf("%s%d", "B", tableStartRow), fmt.Sprintf("%s%d", "B", tableEndRow), tableOPSStyle)
	xlsxFile.SetCellStyle(sheet, fmt.Sprintf("%s%d", "A", tableStartRow), fmt.Sprintf("%s%d", "D", tableStartRow), tableHeaderStyle)
	xlsxFile.SetCellStyle(sheet, fmt.Sprintf("%s%d", "A", tableEndRow), fmt.Sprintf("%s%d", "D", tableEndRow), tableFinalStyle)

	err = xlsxFile.SaveAs(output)
	if err != nil {
		log.Printf("error: %v", err)
	}
}
