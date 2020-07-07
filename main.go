package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/gin-gonic/gin"
)

func main() {
	exe, _ := os.Executable()    // 実行ファイルのフルパス
	rootDir := filepath.Dir(exe) // 実行ファイルのあるディレクトリ

	r := gin.Default()
	r.Static("/results", "./results") // 静的ディレクトリとしておかないとHTMLのダウンロードリンクからアクセスできない
	r.LoadHTMLGlob("html/**/*.tmpl")

	// アクセスされたらこれを表示
	r.GET("/", func(c *gin.Context) {
		c.HTML(http.StatusOK, "html/index.tmpl", gin.H{
			"title": "PDF Page Counter",
		})
	})

	// uploadされたらこれ
	r.POST("/", func(c *gin.Context) {
		zipFile, err := c.FormFile("upload")
		if err != nil {
			c.String(http.StatusBadRequest, fmt.Sprintf("get form err: %s", err.Error()))
			return
		}
		log.Println(zipFile.Filename)

		// 特定のディレクトリにファイルをアップロードする
		dst := rootDir + "\\uploaded" + "\\" + filepath.Base(zipFile.Filename)
		log.Println(dst)
		if err := c.SaveUploadedFile(zipFile, dst); err != nil {
			c.String(http.StatusBadRequest, fmt.Sprintf("upload file err: %s", err.Error()))
			return
		}

		// 時刻オブジェクト
		t := time.Now()
		const layout = "2006-01-02_15-04-05"
		tFormat := t.Format(layout)

		// 結果が記載されるcsvのファイル名
		resultFile := tFormat + ".csv"
		resultFile = "results\\" + resultFile

		// outフォルダを削除する
		if err := os.RemoveAll("out"); err != nil {
			fmt.Println(err)
		}

		// outフォルダを作る
		if err := os.Mkdir("out", 0777); err != nil {
			fmt.Println(err)
		}

		// resultFileを削除する
		//if err := os.Remove(resultFile); err != nil {
		//	fmt.Println(err)
		//}

		// unzipする
		out, err3 := exec.Command("7z.exe", "x", "-y", "-o"+rootDir+"\\out", dst).CombinedOutput()
		log.Println("7z.exe", "x", "-y", "-o"+rootDir+"\\out", dst)
		if err3 != nil {
			fmt.Println("7zip command Exec Error")
		}
		fmt.Printf("ls result: \n%s", string(out))

		// resultFileを作成してオープンする
		csvFile, err := os.OpenFile(resultFile, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
		if err != nil {
			log.Fatal(err)
		}
		defer csvFile.Close()

		// resultFileのヘッダー
		fmt.Fprintln(csvFile, "File Name"+","+"Page Count")

		// 再帰でPDFを処理する
		paths := dirwalk(rootDir + `\out`)

		counter := 1
		flag := 0
		for _, path := range paths {
			ext := filepath.Ext(path) // ファイルの拡張子を得る
			if ext == ".pdf" {
				flag++
				log.Println("Processing... " + path)

				// pdfinfoコマンドの出力をゲットする
				pdfinfoOut, err := exec.Command("pdfinfo.exe", path).CombinedOutput()
				if err != nil {
					fmt.Println("pdfinfo command Exec Error")
				}

				s := string(pdfinfoOut)
				sArray := strings.Split(s, "\n") // 改行でスプリットして配列にプッシュ
				var page string
				for _, s := range sArray {
					// ページカウントのパターンにマッチさせる
					if regexp.MustCompile(`Pages:(\s)+?(\d+)`).MatchString(s) == true {
						re := regexp.MustCompile(`Pages:(\s)+?(\d+)`) // 正規表現をコンパイル
						page = re.ReplaceAllString(s, "$2")           // ページ数のみに置換
						//log.Println(page)
					}
				}

				// pathから不要な文字を削除する
				replacedPath := strings.Replace(path, rootDir+"\\out", "", 1)
				//log.Println("Replaced path: " + replacedPath)

				// csvに書き込み
				fmt.Fprintln(csvFile, replacedPath + "," + page)
				counter++

			}
		}
		fmt.Fprintln(csvFile)                                   // 空行
		counterStr := strconv.Itoa(counter)                     // intからstrに変換
		fmt.Fprintln(csvFile, "Total,=SUM(B2:B"+counterStr+")") // ページ合計
		csvFile.Close()

		// nkfコマンドでBOM付きにする
		errNkf := exec.Command("nkf.exe", "-w8", "--overwrite", rootDir+"\\"+resultFile).Run()
		log.Println("nkf.exe", "-w8", "--overwrite", rootDir+"\\"+resultFile)
		if errNkf != nil {
			fmt.Println("nkf command Exec Error")
		}

		if flag == 0 {
			// pdfがなかった場合はこれを返す
			c.String(http.StatusOK, "There is no pdf in the uploaded zip.")
		} else {
			// pdfがあった場合はcsvを返す

			// ※これは挙動がおかしい。Chromeだとファイル名が「ダウンロード」になる。FireFoxだと適当な名前になってエクセルになっちゃう。
			//c.File(resultFile)

			// index.tmplを書き換えて、HTMLからダウンロードさせる
			c.HTML(http.StatusOK, "html/index.tmpl", gin.H{
				"title":           "PDF Page Counter",
				"downloadMessage": "Please download the csv file: ",
				"downloadfile":    tFormat + ".csv",
			})
		}
	})

	r.Run(":8")
}

// 再帰
func dirwalk(dir string) []string {
	files, err := ioutil.ReadDir(dir)
	if err != nil {
		panic(err)
	}

	var paths []string
	for _, file := range files {
		if file.IsDir() {
			paths = append(paths, dirwalk(filepath.Join(dir, file.Name()))...)
			continue
		}
		paths = append(paths, filepath.Join(dir, file.Name()))
	}

	return paths
}

func newCsvWriter(w io.Writer, bom bool) *csv.Writer {
	bw := bufio.NewWriter(w)
	if bom {
		bw.Write([]byte{0xEF, 0xBB, 0xBF})
	}
	return csv.NewWriter(bw)
}
