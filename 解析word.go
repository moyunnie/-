package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"github.com/google/uuid"
	"html/template"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"
)

//	type Document struct {
//		XMLName xml.Name `xml:"document"`
//		Body    Body     `xml:"body"`
//	}
//
//	type Body struct {
//		Paragraphs []Paragraph `xml:"p"`
//	}
//
//	type Paragraph struct {
//		Run Run `xml:"r"`
//	}
//
//	type Run struct {
//		Text []Text `xml:"t"`
//	}
//
// // 定义 XML 解析结构体
//
//	type Text struct {
//		Value string `xml:",chardata"`
//	}
type Document struct {
	XMLName xml.Name `xml:"document"`

	Body Body `xml:"body"`
}

type Body struct {
	XMLName xml.Name `xml:"body"`

	Paragraphs []Paragraph `xml:"p"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"p"`

	Runs []Run `xml:"r"`
}

type Run struct {
	XMLName xml.Name `xml:"r"`

	Text string `xml:"t"`

	Drawing Drawing `xml:"drawing"`
}

type Drawing struct {
	XMLName xml.Name `xml:"drawing"`

	Inline Inline `xml:"inline"`
}

type Inline struct {
	XMLName xml.Name `xml:"inline"`

	Extent Extent `xml:"extent"`

	DocPr DocPr `xml:"docPr"`

	Graphic Graphic `xml:"graphic"`
}

type Extent struct {
	XMLName xml.Name `xml:"extent"`

	Cx int `xml:"cx,attr"`

	Cy int `xml:"cy,attr"`
}

type DocPr struct {
	XMLName xml.Name `xml:"docPr"`

	Id int `xml:"id,attr"`

	Name string `xml:"name,attr"`

	Descr string `xml:"descr,attr"`
}

type Graphic struct {
	XMLName xml.Name `xml:"graphic"`

	GraphicData GraphicData `xml:"graphicData"`
}

type GraphicData struct {
	XMLName xml.Name `xml:"graphicData"`

	Pic Pic `xml:"pic"`
}

type Pic struct {
	XMLName xml.Name `xml:"pic"`

	NvPicPr NvPicPr `xml:"nvPicPr"`

	BlipFill BlipFill `xml:"blipFill"`

	SpPr SpPr `xml:"spPr"`
}

type NvPicPr struct {
	XMLName xml.Name `xml:"nvPicPr"`

	CNvPr CNvPr `xml:"cNvPr"`

	CNvPicPr CNvPicPr `xml:"cNvPicPr"`
}

type CNvPr struct {
	XMLName xml.Name `xml:"cNvPr"`

	Id int `xml:"id,attr"`

	Name string `xml:"name,attr"`

	Descr string `xml:"descr,attr"`
}

type CNvPicPr struct {
	XMLName xml.Name `xml:"cNvPicPr"`

	PicLocks PicLocks `xml:"picLocks"`
}

type PicLocks struct {
	XMLName xml.Name `xml:"picLocks"`

	NoChangeAspect int `xml:"noChangeAspect,attr"`
}

type BlipFill struct {
	XMLName xml.Name `xml:"blipFill"`

	Blip Blip `xml:"blip"`
}

type Blip struct {
	XMLName xml.Name `xml:"blip"`

	Embed string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
}

type SpPr struct {
	XMLName xml.Name `xml:"spPr"`

	Xfrm Xfrm `xml:"xfrm"`

	PrstGeom PrstGeom `xml:"prstGeom"`
}

type Xfrm struct {
	XMLName xml.Name `xml:"xfrm"`

	Off Off `xml:"off"`

	Ext Ext `xml:"ext"`
}

type Off struct {
	XMLName xml.Name `xml:"off"`

	X int `xml:"x,attr"`

	Y int `xml:"y,attr"`
}

type Ext struct {
	XMLName xml.Name `xml:"ext"`

	Cx int `xml:"cx,attr"`

	Cy int `xml:"cy,attr"`
}

type PrstGeom struct {
	XMLName xml.Name `xml:"prstGeom"`

	Prst string `xml:"prst,attr"`

	AvLst AvLst `xml:"avLst"`
}

type AvLst struct {
	XMLName xml.Name `xml:"avLst"`
}

func Zip(out string, input string) {
	dst := out
	archive, err := zip.OpenReader(input)
	if err != nil {
		panic(err)
	}
	defer archive.Close()

	for _, f := range archive.File {
		filePath := filepath.Join(dst, f.Name)
		fmt.Println("unzipping file ", filePath)

		if !strings.HasPrefix(filePath, filepath.Clean(dst)+string(os.PathSeparator)) {
			fmt.Println("invalid file path")
			return
		}
		if f.FileInfo().IsDir() {
			fmt.Println("creating directory...")
			os.MkdirAll(filePath, os.ModePerm)
			continue
		}

		if err := os.MkdirAll(filepath.Dir(filePath), os.ModePerm); err != nil {
			panic(err)
		}

		dstFile, err := os.OpenFile(filePath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
		if err != nil {
			panic(err)
		}

		fileInArchive, err := f.Open()
		if err != nil {
			panic(err)
		}

		if _, err := io.Copy(dstFile, fileInArchive); err != nil {
			panic(err)
		}

		dstFile.Close()
		fileInArchive.Close()
	}
}

func main() {
	uuids, _ := uuid.NewUUID()
	fmt.Println(uuids)
	os.Rename("3213.docx", fmt.Sprintf("%s.zip", uuids))
	outnt := "ceshi12"
	Zip(outnt, fmt.Sprintf("%s.zip", uuids))

	//os.Rename()
	// 读取 Word 文档的 XML 内容
	xmlContent, err := ioutil.ReadFile(fmt.Sprintf("%s/word/document.xml", outnt))
	if err != nil {
		log.Fatal(err)
	}
	// 解析 XML 内容
	var doc Document
	err = xml.Unmarshal(xmlContent, &doc)
	if err != nil {
		log.Fatal(err)
	}
	// 生成富文本内容
	var sb strings.Builder
	for _, p := range doc.Body.Paragraphs {
		for _, r := range p.Runs {
			if r.Text != "" {
				sb.WriteString(template.HTMLEscapeString(r.Text))
				sb.WriteString("<br>")

			}
			if r.Drawing.Inline.DocPr.Name != "" {
				imgPath := fmt.Sprintf("%s/word/media/image%d.png", outnt, r.Drawing.Inline.DocPr.Id)
				//imgData, err := ioutil.ReadFile(imgPath)
				//if err != nil {
				//	log.Fatal(err)
				//}
				//imgSrc := fmt.Sprintf("data:image/png;base64,%s", template.HTMLEscapeString(base64.StdEncoding.EncodeToString(imgData)))
				sb.WriteString(fmt.Sprintf(`<img src="%s" alt="%s" />`, imgPath, r.Drawing.Inline.DocPr.Descr))
				sb.WriteString("<br>")
			}


			
		}
		sb.WriteString("<br>")
	}
	htmlContent := sb.String()
	fmt.Println(htmlContent)
}
