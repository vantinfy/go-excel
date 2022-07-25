package main

import (
	"archive/zip"
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

func main() {
	fileName := ``
	fmt.Println("将要处理的excel文件放在与本程序同目录下 ")
	fmt.Println("复制文件名(包括文件拓展名) 输入后回车")

	reader := bufio.NewReader(os.Stdin) // 标准输入输出
	msg, _ := reader.ReadString('\n')   // 回车结束
	fileName = strings.TrimSpace(msg)   // 去除最后一个空格

	fmt.Println("复制xls源文件并压缩...")
	_, err := CopyFile(fileName, "./tmp.zip")
	if err != nil {
		fmt.Println("复制出错", err)
		return
	}

	fmt.Println("解压文件...")
	UnzipFile("./tmp.zip", "./tmp")
	_ = os.Remove("./tmp.zip") // 解压完成后删除临时文件

	fmt.Println("读取xls文件...")
	maps, err := ReadExcel(fileName, "Sheet1")
	if err != nil {
		fmt.Println("读取xls文件出错", err)
		return
	}

	// 最后图片的存储路径
	fmt.Println("重命名文件...")
	HandleXML(maps, "./output/")
	fmt.Println("重名命完成")

	fmt.Println("删除临时文件夹...")
	err = os.RemoveAll("./tmp")
	if err != nil {
		fmt.Println("删除临时文件夹出错", err)
	}

	fmt.Println("按任意键回车退出")
	_, _ = fmt.Scan(&fileName)
	os.Exit(1)
}

func HandleXML(maps map[string]string, output string) {
	xmlBytes, err := os.ReadFile("./tmp/xl/cellimages.xml")
	if err != nil {
		fmt.Println("读取xml文件失败", err)
		return
	}

	err = os.Mkdir(output, 0777)
	if err != nil {
		fmt.Println("创建输出文件夹失败", err)
		return
	}

	// 暂时先用正则匹配吧
	reg := regexp.MustCompile(`<xdr:nvPicPr><xdr:cNvPr id=".+?" name="ID_.+? r:embed="rId.+?"/>`)
	matches := reg.FindAllString(string(xmlBytes), -1)
	prefix := "./tmp/xl/media/image"
	suffix := ".png"
	for _, match := range matches {
		//fmt.Println(match)
		nameAndEmbed := strings.Split(match, `name="`)
		nameSplitEmbed := strings.Split(nameAndEmbed[1], `"/>`) // name nameSplitEmbed[0]

		// 新文件名
		fileName := output + maps[strings.Split(nameSplitEmbed[0], `"`)[0]] + suffix
		imgNo := strings.Split(nameSplitEmbed[len(nameSplitEmbed)-2], `r:embed="rId`) // id imgNo[len-1]
		// 旧文件名
		oldName := prefix + imgNo[len(imgNo)-1]
		err = os.Rename(oldName+suffix, fileName)
		if err != nil {
			jpegSuffix := ".jpeg"
			err = os.Rename(oldName+jpegSuffix, fileName)
			if err != nil {
				fmt.Println("重命名错误", err)
			}
		}
	}

	// todo xml处理 样例: <xdr:nvPicPr><xdr:cNvPr id=".+?" name="ID_.+? r:embed="rId.+?"/>
	//cellImages := CellImages{}
	//err = xml.Unmarshal(xmlBytes, &cellImages)
	//if err != nil {
	//	fmt.Println("反序列化错误")
	//	return
	//}
	//fmt.Println(len(cellImages.CellImage), cellImages)
}

type CellImages struct {
	CellImage []CellImage `xml:"cellImage" json:"cellImage"`
}

type CellImage struct {
	Pic `xml:"pic" json:"pic"`
}

type Pic struct {
	NvPicPr `json:"nvPicPr" json:"nvPicPr"`
}

type NvPicPr struct {
	CNvPr `xml:"cNvPr" json:"cNvPr"`
}

type CNvPr struct {
	Id   string `xml:"id" json:"id"`
	Name string `xml:"name" json:"name"`
}

func ReadExcel(file string, sheet string) (map[string]string, error) {
	mapping := map[string]string{} // cellImg-name

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println("文件打开错误", err)
		return nil, err
	}

	rows := xlsx.GetRows(sheet)
	for _, row := range rows {
		//fmt.Println("rows", row)
		for _, s := range row {
			if strings.Contains(s, "=DISPIMG") {
				res := strings.Split(s, `"`)
				mapping[res[1]] = row[0]
			}
		}
	}
	return mapping, nil
}

func CopyFile(source, destination string) (int64, error) {
	// 要复制的文件是否存在
	sourceFileStat, err := os.Stat(source)
	if err != nil {
		return 0, err
	}

	// 是否为常规文件
	if !sourceFileStat.Mode().IsRegular() {
		return 0, fmt.Errorf("%s is not a regular file", source)
	}

	// 打开源文件
	src, err := os.Open(source)
	if err != nil {
		fmt.Println("源文件不存在", src)
		return 0, err
	}
	defer func() {
		_ = src.Close()
	}()

	// 目标文件
	dst, err := os.Create(destination)
	if err != nil {
		fmt.Println("创建文件副本失败", err)
		return 0, err
	}
	defer func() {
		_ = dst.Close()
	}()

	return io.Copy(dst, src)
}

func UnzipFile(fileName, dst string) {
	archive, err := zip.OpenReader(fileName)
	if err != nil {
		fmt.Println("打开文件错误", err)
		return
	}
	defer func() {
		_ = archive.Close()
		r := recover()
		if r != nil {
			fmt.Println("解压过程出现错误", r)
		}
	}()

	for _, f := range archive.File {
		filePath := filepath.Join(dst, f.Name)
		//fmt.Println("unzipping file ", filePath)

		if !strings.HasPrefix(filePath, filepath.Clean(dst)+string(os.PathSeparator)) {
			fmt.Println("文件路径非法")
			return
		}
		if f.FileInfo().IsDir() {
			//fmt.Println("creating directory...")
			_ = os.MkdirAll(filePath, os.ModePerm)
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

		_ = dstFile.Close()
		_ = fileInArchive.Close()
	}
}
