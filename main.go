package main

import (
	"bytes"
	"crypto"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/maps"
	"golang.org/x/exp/slices"
)

func main() {
	payloadSign := getSignature("23y0ufFl5YxIyGrI8hWRUZmKkvtSjLQA", "123456789abcdefg")
	headerSign := getSignature("fTN2pfuisxTavbTuYVSsNJHetwq5bJvC", "QkjjtiLM2dCratiA")

	p := payload{
		AppId:           "NcApplication",
		Key:             "3C502C97ABDA40D0A60FBEE50FAAD1DA",
		NonceHeader:     "123456789abcdefg",
		PassHeader:      "zdww",
		SignatureHeader: payloadSign.sign,
		TimestampHeader: payloadSign.timestamp,
	}
	d, err := json.Marshal(p)
	if err != nil {
		log.Fatal(err)
	}

	url := "http://bmfw.www.gov.cn/bjww/interface/interfaceJson"
	req, err := http.NewRequest("POST", url, bytes.NewBuffer(d))
	if err != nil {
		panic(err)
	}

	req.Header = make(http.Header)
	req.Header.Add("Accept", "application/json, text/javascript, */*; q=0.01")
	req.Header.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8")
	req.Header.Add("Referer", "http:/ /bmfw.www.gov.cn/")
	req.Header.Add("Origin", "http://bmfw.www.gov.cn")
	req.Header.Add("Host", "bmfw.www.gov.cn")
	req.Header.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36")
	req.Header.Add("Content-Type", "application/json; charset=UTF-8")
	req.Header.Add("x-wif-nonce", "QkjjtiLM2dCratiA")
	req.Header.Add("x-wif-paasid", "smt-application")
	req.Header.Add("x-wif-signature", headerSign.sign)
	req.Header.Add("x-wif-timestamp", headerSign.timestamp)

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		panic(err)
	}

	if resp.StatusCode != http.StatusOK {
		fmt.Println("未知响应状态，请联系开发者")
		return
	}

	var result Resp
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		panic(err)
	}

	if err := writeToxlsx(result.Data); err != nil {
		panic(err)
	}
}

func writeToxlsx(data areaData) error {
	f := excelize.NewFile()
	f.NewSheet("Sheet1")

	f.SetCellValue("Sheet1", "A1", "国内疫情风险地区")
	f.SetCellValue("Sheet1", "A2", fmt.Sprintf("最近更新时间：%s", data.UpdatedTime))
	f.SetCellValue("Sheet1", "A3", fmt.Sprintf("高风险地区：%d, 低风险地区：%d", data.HightRiskCnt, data.LowRiskCnt))

	f.SetCellValue("Sheet1", "A5", "风险等级")
	f.SetCellValue("Sheet1", "B5", "省份")
	f.SetCellValue("Sheet1", "C5", "城市")
	f.SetCellValue("Sheet1", "D5", "区县")
	f.SetCellValue("Sheet1", "E5", "区域")

	setCell := func(sheet string, cell []string, value []string) error {
		for idx := range cell {
			if err := f.SetCellValue(sheet, cell[idx], value[idx]); err != nil {
				return err
			}
		}

		return nil
	}

	for idx, a := range data.HightAreas {
		lineNum := idx + 6
		if err := setCell(
			"Sheet1",
			[]string{
				fmt.Sprintf("A%d", lineNum),
				fmt.Sprintf("B%d", lineNum),
				fmt.Sprintf("C%d", lineNum),
				fmt.Sprintf("D%d", lineNum),
				fmt.Sprintf("E%d", lineNum),
			},
			[]string{
				"高",
				a.Province,
				a.City,
				a.County,
				strings.Join(a.Communitys, ","),
			},
		); err != nil {
			panic(err)
		}
	}

	for idx, a := range data.LowAreas {
		lineNum := idx + 6 + len(data.HightAreas)
		if err := setCell(
			"Sheet1",
			[]string{
				fmt.Sprintf("A%d", lineNum),
				fmt.Sprintf("B%d", lineNum),
				fmt.Sprintf("C%d", lineNum),
				fmt.Sprintf("D%d", lineNum),
				fmt.Sprintf("E%d", lineNum),
			},
			[]string{
				"低",
				a.Province,
				a.City,
				a.County,
				strings.Join(a.Communitys, ","),
			},
		); err != nil {
			panic(err)
		}
	}

	sheetName := "风险城市整理表格"
	sheet2 := f.NewSheet(sheetName)
	f.SetActiveSheet(sheet2)

	citySet := make(map[string]struct{})
	for _, a := range data.HightAreas {
		city := []rune(a.City)
		province := []rune(a.Province)
		switch {
		case a.City == "杨凌示范区":
			citySet["杨凌"] = struct{}{}
		case a.City == "临夏回族自治州":
			citySet["临夏"] = struct{}{}
		case a.City == "省直辖县级行政单位":
			citySet[a.County] = struct{}{}
		case city[len(city)-1] == []rune("区")[0] && province[len(province)-1] == []rune("市")[0]:
			a.City = a.Province
			fallthrough
		case city[len(city)-1] == []rune("市")[0]:
			citySet[strings.Replace(a.City, "市", "", 1)] = struct{}{}
		default:
			citySet[a.City] = struct{}{}
		}
	}

	cities := maps.Keys(citySet)
	slices.Sort(cities)
	f.SetCellValue(sheetName, "A1", "高风险")
	for idx, c := range cities {
		if err := f.SetCellValue(sheetName, fmt.Sprintf("A%d", idx+2), c); err != nil {
			panic(err)
		}
	}

	return f.SaveAs("国内疫情风险等级数据.xlsx")
}

type Resp struct {
	Code int      `json:"code"`
	Msg  string   `json:"msg"`
	Data areaData `json:"data"`
}

type areaData struct {
	UpdatedTime  string `json:"end_update_time"`
	HightRiskCnt int    `json:"hcount"`
	LowRiskCnt   int    `json:"lcount"`
	HightAreas   []area `json:"highlist"`
	LowAreas     []area `json:"lowlist"`
}

type area struct {
	Type       string   `json:"type"`
	Province   string   `json:"province"`
	City       string   `json:"city"`
	County     string   `json:"county"`
	AreaName   string   `json:"area_name"`
	Communitys []string `json:"communitys"`
}

type payload struct {
	AppId           string `json:"appId"`
	Key             string `json:"key"`
	NonceHeader     string `json:"nonceHeader"`
	PassHeader      string `json:"paasHeader"`
	SignatureHeader string `json:"signatureHeader"`
	TimestampHeader string `json:"timestampHeader"`
}

type signature struct {
	sign      string
	timestamp string
}

func getSignature(token, nonce string) signature {
	timestamp := time.Now().UnixMilli() / 1000
	hash := crypto.SHA256.New()
	hash.Write([]byte(fmt.Sprintf("%d", timestamp)))
	hash.Write([]byte(token))
	hash.Write([]byte(nonce))
	hash.Write([]byte(fmt.Sprintf("%d", timestamp)))

	return signature{
		sign:      strings.ToUpper(hex.EncodeToString(hash.Sum(nil))),
		timestamp: fmt.Sprintf("%d", timestamp),
	}
}
