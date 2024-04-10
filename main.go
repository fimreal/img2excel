package main

import (
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/google/uuid"
)

const (
	YOUDAO_URL = "https://openapi.youdao.com/ocr_table"
	APP_KEY    = ""
	APP_SECRET = ""
)

func truncate(q string) string {
	size := len([]rune(q))
	if size <= 20 {
		return q
	}
	return q[:10] + strconv.Itoa(size) + q[size-10:size]
}

func encrypt(signStr string) string {
	hash := sha256.New()
	hash.Write([]byte(signStr))
	return hex.EncodeToString(hash.Sum(nil))
}

func doRequest(data map[string]string) []byte {
	client := &http.Client{}
	formData := url.Values{}
	for key, value := range data {
		formData.Set(key, value)
	}
	req, err := http.NewRequest("POST", YOUDAO_URL, strings.NewReader(formData.Encode()))
	if err != nil {
		fmt.Println("Error creating request:", err)
		return nil
	}

	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	resp, err := client.Do(req)
	if err != nil {
		fmt.Println("Error making request:", err)
		return nil
	}

	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		fmt.Println("Error reading response:", err)
		return nil
	}

	return body
}

func tansToExcel(response []byte) []byte {
	// 解析 JSON 数据
	var data map[string]interface{}
	err := json.Unmarshal([]byte(response), &data)
	if err != nil {
		fmt.Println("Error decoding JSON:", err)
		return nil
	}
	resultData, ok := data["Result"].(map[string]interface{})
	if !ok {
		fmt.Println("Result field not found or not a map")
		return nil
	}
	resultValue, ok := resultData["tables"]
	if !ok {
		fmt.Println("tables field not found in Result")
		return nil
	}
	tables, ok := resultValue.([]interface{})
	if !ok {
		fmt.Println("tables field is not an array")
		return nil
	}
	if len(tables) == 0 {
		fmt.Println("tables array is empty")
		return nil
	}
	firstTable := tables[0].(string)
	decodedData, err := base64.StdEncoding.DecodeString(firstTable)
	if err != nil {
		fmt.Println("Error decoding Base64:", err)
		return nil
	}

	return decodedData
}

func ocr(filename string) []byte {
	f, err := os.ReadFile(filename) // 读取文件内容
	if err != nil {
		fmt.Println("Error reading file:", err)
		return nil
	}
	q := base64.StdEncoding.EncodeToString(f)

	data := make(map[string]string)
	data["type"] = "1"
	data["q"] = q
	data["docType"] = "excel"
	data["signType"] = "v3"
	curtime := strconv.FormatInt(time.Now().Unix(), 10)
	data["curtime"] = curtime
	salt := uuid.New().String()
	signStr := APP_KEY + truncate(q) + salt + curtime + APP_SECRET
	sign := encrypt(signStr)
	data["appKey"] = APP_KEY
	data["salt"] = salt
	data["sign"] = sign

	response := doRequest(data)
	if response != nil {
		fmt.Println(string(response))
	}
	return tansToExcel(response)

}

func main() {
	ExcelFileContent := ocr(os.Args[1])
	os.WriteFile(os.Args[1]+".xlsx", ExcelFileContent, 0644)
}
