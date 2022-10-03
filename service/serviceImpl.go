package service

import (
	"city/model"
	"context"
	"crypto/tls"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"net/smtp"
	"os"
	"time"

	asposecellscloud "github.com/aspose-cells-cloud/aspose-cells-cloud-go/v22"
	pdf "github.com/balacode/one-file-pdf"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/bson/primitive"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
	gomail "gopkg.in/mail.v2"
)

type Connection struct {
	Server       string
	Database     string
	Collection   string
	Colllection2 string
}

var Collection *mongo.Collection
var CategoryCollection *mongo.Collection
var ctx = context.TODO()
var insertDocs int

func (e *Connection) Connect() {
	clientOptions := options.Client().ApplyURI(e.Server)
	client, err := mongo.Connect(ctx, clientOptions)
	if err != nil {
		log.Fatal(err)
	}

	err = client.Ping(ctx, nil)
	if err != nil {
		log.Fatal(err)
	}

	Collection = client.Database(e.Database).Collection(e.Collection)
	CategoryCollection = client.Database(e.Database).Collection(e.Colllection2)
}

func (e *Connection) InsertAllData(cityData []model.CityData, field string) (int, error) {

	data, err := e.SearchDataInCategories(field)
	if err != nil {
		return 0, err
	}
	id := data[0].ID

	for i := range cityData {
		cityData[i].CategoriesId = id
		_, err := Collection.InsertOne(ctx, cityData[i])

		if err != nil {
			return 0, errors.New("Unable To Insert New Record")
		}
		insertDocs = i + 1
	}
	return insertDocs, nil
}

func (e *Connection) DeleteData(cityData string) (string, error) {

	id, err := primitive.ObjectIDFromHex(cityData)

	if err != nil {
		return "", err
	}

	filter := bson.D{primitive.E{Key: "_id", Value: id}}

	cur, err := Collection.DeleteOne(ctx, filter)

	if err != nil {
		return "", err
	}

	if cur.DeletedCount == 0 {
		return "", errors.New("Unable To Delete Data")
	}

	return "Deleted Successfully", nil
}

func (e *Connection) SearchData(searchBoth model.SearchBoth) ([]byte, string, error) {
	var data []*model.CityData
	var cursor *mongo.Cursor
	var err error
	var dataty []byte
	os.MkdirAll("data/download", os.ModePerm)
	dir := "data/download/"
	file := "searchResult" + fmt.Sprintf("%v", time.Now().Format("3_4_5_pm"))
	csvFile, err := os.Create(dir + file + ".csv")
	str := "please provide value of either city or category"
	if (searchBoth.City != "") && (searchBoth.Category != "") {
		categoryData, error := e.SearchDataInCategories(searchBoth.Category)

		if error != nil {
			return dataty, file, err
		}

		id := categoryData[0].ID
		cursor, err = Collection.Find(ctx, bson.D{primitive.E{Key: "categories_id", Value: id}, primitive.E{Key: "city", Value: searchBoth.City}})

		if err != nil {
			return dataty, file, err
		}
		str = "No data present in city data db for given category or city"
	} else if searchBoth.City != "" {
		cursor, err = Collection.Find(ctx, bson.D{primitive.E{Key: "city", Value: searchBoth.City}})

		if err != nil {
			return dataty, file, err
		}
		str = "No data present in db for given city name"
	} else if searchBoth.Category != "" {
		categoryData, error := e.SearchDataInCategories(searchBoth.Category)

		if error != nil {
			return dataty, file, err
		}

		id := categoryData[0].ID
		cursor, err = Collection.Find(ctx, bson.D{primitive.E{Key: "categories_id", Value: id}})

		if err != nil {
			return dataty, file, err
		}
		str = "No data present in city data db for given category"
	}

	for cursor.Next(ctx) {
		var e model.CityData
		err := cursor.Decode(&e)
		if err != nil {
			return dataty, file, err
		}
		data = append(data, &e)
	}

	if data == nil {
		return dataty, file, errors.New(str)
	}

	if err != nil {
		fmt.Println(err)
	}
	defer csvFile.Close()
	writer := csv.NewWriter(csvFile)

	header := []string{"ID", "Title", "Name", "Address", "Latitude", "Longitude", "Website", "ContactNumber", "User", "City", "Country", "PinCode", "UpdatedBy", "CategoriesId"}
	if err := writer.Write(header); err != nil {
		return dataty, file, err
	}

	for _, r := range data {
		var csvRow []string
		csvRow = append(csvRow, fmt.Sprintf("%v", r.ID), r.Title, r.Name, r.Address, fmt.Sprintf("%f", r.Latitude), fmt.Sprintf("%f", r.Longitude), r.Website, fmt.Sprintf("%v", r.ContactNumber),
			r.User, r.City, r.Country, fmt.Sprintf("%v", r.PinCode), r.UpdatedBy, fmt.Sprintf("%v", r.CategoriesId))
		if err := writer.Write(csvRow); err != nil {
			return dataty, file, err
		}
	}

	// remember to flush!
	writer.Flush()

	dataty, err = ioutil.ReadFile(dir + file + ".csv")
	if err != nil {
		log.Fatal(err)
	}

	res, err := WriteIntoPDF2(data)
	//res2, err := WriteIntoPDF3(data)
	//emailStar()
	emailStar2()
	fmt.Println(res)
	//writeIntoPDF(file, dir)

	return dataty, file, nil
}

func (e *Connection) SearchDataByKeyAndValue(reqBody model.Search) ([]*model.CityData, error) {
	var data []*model.CityData

	cursor, err := Collection.Find(ctx, bson.D{primitive.E{Key: reqBody.Key, Value: reqBody.Value}})

	if err != nil {
		return data, err
	}

	for cursor.Next(ctx) {
		var e model.CityData
		err := cursor.Decode(&e)
		if err != nil {
			return data, err
		}
		data = append(data, &e)
	}

	if data == nil {
		return data, errors.New("No data present in db for given city name")
	}

	return data, nil
}

func (e *Connection) UpdateData(cityData model.CityData, field string) (string, error) {

	id, err := primitive.ObjectIDFromHex(field)

	if err != nil {
		return "", err
	}

	filter := bson.D{primitive.E{Key: "_id", Value: id}}

	update := bson.D{primitive.E{Key: "$set", Value: cityData}}

	err2 := Collection.FindOneAndUpdate(ctx, filter, update).Decode(e)

	if err2 != nil {
		return "", err2
	}
	return "Data Updated Successfully", nil
}

func (e *Connection) InsertAllDataInCategories(categoryData []model.Categories) (int, error) {
	for i := range categoryData {
		_, err := CategoryCollection.InsertOne(ctx, categoryData[i])

		if err != nil {
			return 0, errors.New("Unable To Insert New Record")
		}
		insertDocs = i + 1
	}
	return insertDocs, nil
}

func (e *Connection) DeleteDataInCategories(categoryId string) (string, error) {

	id, err := primitive.ObjectIDFromHex(categoryId)

	if err != nil {
		return "", err
	}

	filter := bson.D{primitive.E{Key: "_id", Value: id}}

	cur, err := CategoryCollection.DeleteOne(ctx, filter)

	if err != nil {
		return "", err
	}

	if cur.DeletedCount == 0 {
		return "", errors.New("Unable To Delete Data")
	}

	return "Deleted Successfully", nil
}

func (e *Connection) SearchDataInCategories(name string) ([]*model.Categories, error) {
	var data []*model.Categories

	cursor, err := CategoryCollection.Find(ctx, bson.D{primitive.E{Key: "category", Value: name}})

	if err != nil {
		return data, err
	}

	for cursor.Next(ctx) {
		var e model.Categories
		err := cursor.Decode(&e)
		if err != nil {
			return data, err
		}
		data = append(data, &e)
	}

	if data == nil {
		return data, errors.New("No data present in db for given category")
	}
	return data, nil
}

func (e *Connection) UpdateDataInCategories(cityData model.Categories, field string) (string, error) {

	id, err := primitive.ObjectIDFromHex(field)

	if err != nil {
		return "", err
	}

	filter := bson.D{primitive.E{Key: "_id", Value: id}}

	update := bson.D{primitive.E{Key: "$set", Value: cityData}}

	err2 := CategoryCollection.FindOneAndUpdate(ctx, filter, update).Decode(e)

	if err2 != nil {
		return "", err2
	}
	return "Data Updated Successfully", nil
}

func writeIntoPDF(file, dir string) {

	instance := asposecellscloud.NewCellsApiService(os.Getenv("ProductClientId"), os.Getenv("ProductClientSecret"))
	files, err := os.Open(dir + file + ".csv")
	if err != nil {
		fmt.Println(err)
		return
	}
	convertWorkbookOpts := new(asposecellscloud.CellsWorkbookPutConvertWorkbookOpts)

	convertWorkbookOpts.Format = "pdf"

	value, response, err1 := instance.CellsWorkbookPutConvertWorkbook(files, convertWorkbookOpts)
	fmt.Println(response)
	if err1 != nil {
		fmt.Println(err1)
		return

	}

	file1, err2 := os.Create(dir + file + "_" + ".pdf")

	if err2 != nil {
		fmt.Println(err2)
		return

	}

	if _, err3 := file1.Write(value); err3 != nil {
		fmt.Println(err3)
		return

	}

	file1.Close()

}

func WriteIntoPDF2(converted []*model.CityData) (string, error) {
	fmt.Println(`Generating a "Hello World" PDF...`)

	rep, err := json.Marshal(converted)
	if err != nil {
		log.Fatal(err)
	}
	s := string(rep)
	fmt.Println(s)
	m := make(map[string]string)
	err2 := json.Unmarshal(rep, &m)
	if err != nil {
		log.Fatal(err2)
	}
	for key, val := range m {
		fmt.Printf("%s, %s", key, val)
	}

	// create a new PDF using 'A4' page size
	var pdf = pdf.NewPDF("A4")

	// set the measurement units to centimeters
	pdf.SetUnits("cm")

	// draw a grid to help us align stuff (just a guide, not necessary)
	pdf.DrawUnitGrid()

	// draw the word 'HELLO' in orange, using 100pt bold Helvetica font
	// - text is placed on top of, not below the Y-coordinate
	// - you can use method chaining
	var y float64 = 1.0
	var z float64 = 1.0
	for _, v := range converted {
		pdf.SetFont("Helvetica-Bold", 14).
			//SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "_id : "+v.ID.String())
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Title : "+v.Title)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Name : "+v.Name)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Address : "+v.Address)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Website"+v.Website)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "UpdatedBy : "+v.UpdatedBy)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "User : "+v.User)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "City : "+v.City)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Country : "+v.Country)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "UpdatedBy : "+v.UpdatedBy)
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		pdf.SetFont("Helvetica-Bold", 14).
			//	SetXY(y, z).
			SetColor("black").
			DrawTextAt(y, z, "Categories_Id : "+v.CategoriesId.Hex())
		z++
		if z >= 29 {
			pdf.AddPage()
			z = 0
		}
		z++

	}
	//	DrawTextInBox(1.5, 2.5, 20.5, 10.5, "ok", s)

	// draw the word 'WORLD' in blue-violet, using 100pt Helvetica font
	// note that here we use the colo(u)r hex code instead
	// of its name, using the CSS/HTML format: #RRGGBB
	/*pdf.SetXY(1, 2).
	SetColor("black").
	SetFont("Helvetica", 14).
	DrawText("Welcome to this New WORLD!")*/

	// draw a flower icon using 300pt Zapf-Dingbats font
	/*pdf.SetX(7).SetY(17).
	SetColorRGB(255, 0, 0).
	SetFont("ZapfDingbats", 300).
	DrawText("a")*/

	// save the file:
	// if the file exists, it will be overwritten
	// if the file is in use, prints an error message
	pdf.SaveFile("hello.pdf")
	return "saved", nil
}

func WriteIntoPDF3(converted []*model.CityData) (string, error) {
	fmt.Println(`Generating a "Hello World" PDF...`)

	rep, err := json.Marshal(converted)
	if err != nil {
		log.Fatal(err)
	}
	s := string(rep)
	fmt.Println(s)
	m := make(map[string]string)
	err2 := json.Unmarshal(rep, &m)
	if err != nil {
		log.Fatal(err2)
	}
	for key, val := range m {
		fmt.Printf("%s, %s", key, val)
	}

	// create a new PDF using 'A4' page size
	var pdf = pdf.NewPDF("A4")

	// set the measurement units to centimeters
	pdf.SetUnits("cm")

	// draw a grid to help us align stuff (just a guide, not necessary)
	pdf.DrawUnitGrid()

	// draw the word 'HELLO' in orange, using 100pt bold Helvetica font
	// - text is placed on top of, not below the Y-coordinate
	// - you can use method chaining
	var z float64 = 1.0
	var y float64 = 1.0
	//	for _, v := range converted {
	pdf.SetFont("Helvetica-Bold", 14).
		//SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "_id")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Title")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Name")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Address")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Latitude")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Longitude")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "website")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "Contact_no")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "User")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawBox(z, y, 5, 5, true).
		DrawTextInBox(z, y, 5, 5, "", "city")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	pdf.SetFont("Helvetica-Bold", 14).
		//	SetXY(y, z).
		SetColor("black").
		DrawTextInBox(z, y, 5, 5, "", "Country")
	z++
	if z >= 29 {
		pdf.AddPage()
		z = 0
	}
	z++

	//	}
	//	DrawTextInBox(1.5, 2.5, 20.5, 10.5, "ok", s)

	// draw the word 'WORLD' in blue-violet, using 100pt Helvetica font
	// note that here we use the colo(u)r hex code instead
	// of its name, using the CSS/HTML format: #RRGGBB
	/*pdf.SetXY(1, 2).
	SetColor("black").
	SetFont("Helvetica", 14).
	DrawText("Welcome to this New WORLD!")*/

	// draw a flower icon using 300pt Zapf-Dingbats font
	/*pdf.SetX(7).SetY(17).
	SetColorRGB(255, 0, 0).
	SetFont("ZapfDingbats", 300).
	DrawText("a")*/

	// save the file:
	// if the file exists, it will be overwritten
	// if the file is in use, prints an error message
	pdf.SaveFile("hello.pdf")
	return "saved", nil
}

func emailStar() {

	from := "ranveer.singh@gridinfocom.com"
	password := "xsmtpsib-2af236a5040e4b54343f4fe5b59826e9e7588b2e33d160249ab60a3060fbf348-LAnNRJT0sxHcPqYr"

	toEmailAddress := "anurag.singh@gridinfocom.com"
	to := []string{toEmailAddress}

	host := "smtp-relay.sendinblue.com"
	port := "587"
	address := host + ":" + port
	subject := "Subject: This is the subject of the mail\n"
	body := "Hii Ramashankar this is anurag here"
	message := []byte(subject + body)

	auth := smtp.PlainAuth("", from, password, host)

	err := smtp.SendMail(address, auth, from, to, message)
	if err != nil {
		panic(err)
	}

	fmt.Println("email sent")

}

func emailStar2() {
	m := gomail.NewMessage()

	// Set E-Mail sender
	/*	m.SetHeader("From", "ranveer.singh@gridinfocom.com")

		// Set E-Mail receivers
		m.SetHeader("To", "anurag.singh@gridinfocom.com")

		m.SetHeader("Cc", "ramashankar.kumar@gridinfocom.com")

		// Set E-Mail subject
		m.SetHeader("Subject", "Gomail test subject")*/

	m.SetHeaders(map[string][]string{
		"From":    {m.FormatAddress("ranveer.singh@gridinfocom.com", "Ranveer")},
		"To":      {"anurag.singh@gridinfocom.com"},
		"Cc":      {"ramashankar.kumar@gridinfocom.com", "vidhi.goel@gridinfocom.com", "mukesh.jangir@gridinfocom.com", "mohd.salman@gridinfocom.com"},
		"Subject": {"To review this mail"},
	})

	// Set E-Mail body. You can set plain text or html with text/html
	m.SetBody("text/plain", `Hii Everyone,

                                   This is for testing purpose, please review
								this mail and let me know what more changes did i need to do.
							
							Thanks,
						Anurag kumar singh`)

	// Settings for SMTP server
	d := gomail.NewDialer("smtp-relay.sendinblue.com", 587, "ranveer.singh@gridinfocom.com", "xsmtpsib-2af236a5040e4b54343f4fe5b59826e9e7588b2e33d160249ab60a3060fbf348-LAnNRJT0sxHcPqYr")

	// This is only needed when SSL/TLS certificate is not valid on server.
	// In production this should be set to false.
	d.TLSConfig = &tls.Config{InsecureSkipVerify: true}

	// Now send E-Mail
	if err := d.DialAndSend(m); err != nil {
		fmt.Println(err)
		panic(err)
	}

	return
}
