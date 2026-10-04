package main

import (
	"fmt"
	"io"
	"log"
	"math/rand"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func downloadXLS(url, filename string) error {

	resp, err := http.Get(url)
	if err != nil {
		return fmt.Errorf("XLSのダウンロードエラー: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("XLSのダウンロードに失敗しました。ステータスコード: %d", resp.StatusCode)
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("XLSファイルの作成エラー: %v", err)
	}
	defer file.Close()

	_, err = io.Copy(file, resp.Body)
	if err != nil {
		return fmt.Errorf("XLSデータの書き込みエラー: %v", err)
	}

	return nil
}

func readStockCodesFromXLSX(filename string) ([]string, error) {

	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("XLSXファイルのオープンエラー: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, fmt.Errorf("シートの取得エラー")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("行の取得エラー: %v", err)
	}

	var stockCodes []string

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 2 {
			continue
		}
		code := row[1]
		if code != "" {
			stockCodes = append(stockCodes, code)
		}
	}

	return stockCodes, nil
}

func main() {
	xlsURL := "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx"

	now := time.Now()
	yearMonth := now.Format("2006-01")

	xlsFilename := fmt.Sprintf("data_j_%s.xlsx", yearMonth)

	if _, err := os.Stat(xlsFilename); err == nil {
		fmt.Println("同じ月のファイルが既に存在します。ダウンロードをスキップします。")
	} else {
		files, err := filepath.Glob("data_j_*.xls*")
		if err != nil {
			log.Fatalf("ファイルの検索に失敗しました: %v", err)
		}
		for _, file := range files {
			if !strings.Contains(file, yearMonth) {
				if err := os.Remove(file); err != nil {
					log.Printf("ファイルの削除に失敗しました: %v", err)
				} else {
					fmt.Printf("古いファイルを削除しました: %s\n", file)
				}
			}
		}

		fmt.Println("XLSをダウンロードしています...")
		if err := downloadXLS(xlsURL, xlsFilename); err != nil {
			log.Fatalf("XLSのダウンロードに失敗しました: %v", err)
		}
		fmt.Println("XLSのダウンロードが完了しました。")
	}

	fmt.Println("XLSXから銘柄コードを読み込んでいます...")
	stockCodes, err := readStockCodesFromXLSX(xlsFilename)
	if err != nil {
		log.Fatalf("銘柄コードの読み込みに失敗しました: %v", err)
	}

	r := rand.New(rand.NewSource(time.Now().UnixNano()))
	randomStockCode := stockCodes[r.Intn(len(stockCodes))]

	fmt.Println("ランダムに選ばれた銘柄コード:", randomStockCode)
}
