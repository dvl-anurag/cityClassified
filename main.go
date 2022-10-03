package main

import (
	"bytes"
	"city/model"
	"city/service"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"strconv"
	"strings"
	"time"
)

const tokenIdForAdmin = "admin000000122343456"
const tokenIdForPosters = "poster000000122343456"

var mongoDetails = service.Connection{}

func init() {
	mongoDetails.Server = "mongodb://localhost:27017"
	mongoDetails.Database = "Dummy"
	mongoDetails.Collection = "test"
	mongoDetails.Colllection2 = "categories"

	mongoDetails.Connect()
}

func addData(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	token := r.Header.Get("tokenid")

	admin := token == tokenIdForAdmin
	poster := token == tokenIdForPosters

	if !(admin || poster) {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "POST" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}
	path := r.URL.Path
	segments := strings.Split(path, "/")
	field := segments[len(segments)-1]
	var cityData []model.CityData

	if err := json.NewDecoder(r.Body).Decode(&cityData); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
		return
	}

	if inserted, err := mongoDetails.InsertAllData(cityData, field); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": strconv.Itoa(inserted) + " Record Inserted Successfully",
		})
	}
}

func deleteData(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	token := r.Header.Get("tokenid")

	if token != tokenIdForAdmin {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "DELETE" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var reqBody map[string]string

	if err := json.NewDecoder(r.Body).Decode(&reqBody); err != nil {
		respondWithError(w, http.StatusBadRequest, "Invalid request")
	}

	name := reqBody["id"]

	if deleted, err := mongoDetails.DeleteData(name); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": deleted,
		})
	}
}

func searchByCity(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	if r.Method != "GET" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var searchBoth model.SearchBoth

	if err := json.NewDecoder(r.Body).Decode(&searchBoth); err != nil {
		respondWithError(w, http.StatusBadRequest, "Invalid request")
	}

	if searchData, fileName, err := mongoDetails.SearchData(searchBoth); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		w.Header().Set("Content-Type", "application/octet-stream")
		w.Header().Set("Content-Disposition", "attachment; filename="+fileName+".xlsx")
		w.Header().Set("Content-Transfer-Encoding", "binary")
		http.ServeContent(w, r, "Workbook.xlxs", time.Now(), bytes.NewReader(searchData))
	}
}

func search(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	if r.Method != "GET" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var reqBody model.Search

	if err := json.NewDecoder(r.Body).Decode(&reqBody); err != nil {
		respondWithError(w, http.StatusBadRequest, "Invalid request")
	}

	if searchData, err := mongoDetails.SearchDataByKeyAndValue(reqBody); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, searchData)
	}
}

func updateData(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()
	token := r.Header.Get("tokenid")

	admin := token == tokenIdForAdmin

	if !(admin) {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "PUT" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var cityData model.CityData

	if err := json.NewDecoder(r.Body).Decode(&cityData); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
		return
	}

	path := r.URL.Path
	segments := strings.Split(path, "/")
	field := segments[len(segments)-1]

	if updated, err := mongoDetails.UpdateData(cityData, field); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": updated,
		})
	}
}

func addDataInCategory(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	token := r.Header.Get("tokenid")

	admin := token == tokenIdForAdmin
	poster := token == tokenIdForPosters

	if !(admin || poster) {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "POST" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var data []model.Categories

	if err := json.NewDecoder(r.Body).Decode(&data); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
		return
	}

	if inserted, err := mongoDetails.InsertAllDataInCategories(data); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": strconv.Itoa(inserted) + " Record Inserted Successfully",
		})
	}
}

func deleteDataInCategory(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	token := r.Header.Get("tokenid")

	if token != tokenIdForAdmin {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "DELETE" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var reqBody map[string]string

	if err := json.NewDecoder(r.Body).Decode(&reqBody); err != nil {
		respondWithError(w, http.StatusBadRequest, "Invalid request")
	}

	name := reqBody["id"]

	if deleted, err := mongoDetails.DeleteDataInCategories(name); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": deleted,
		})
	}
}

func searchByCategory(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	if r.Method != "GET" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var reqBody map[string]string

	if err := json.NewDecoder(r.Body).Decode(&reqBody); err != nil {
		respondWithError(w, http.StatusBadRequest, "Invalid request")
	}

	name := reqBody["category"]

	if searchData, err := mongoDetails.SearchDataInCategories(name); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, searchData)
	}
}

func updateDataInCategory(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()
	token := r.Header.Get("tokenid")

	admin := token == tokenIdForAdmin

	if !(admin) {
		respondWithError(w, http.StatusBadRequest, "Unauthorized")
		return
	}

	if r.Method != "PUT" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	var data model.Categories

	if err := json.NewDecoder(r.Body).Decode(&data); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
		return
	}

	path := r.URL.Path
	segments := strings.Split(path, "/")
	field := segments[len(segments)-1]

	if updated, err := mongoDetails.UpdateDataInCategories(data, field); err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, map[string]string{
			"message": updated,
		})
	}
}

func respondWithJson(w http.ResponseWriter, code int, payload interface{}) {
	response, _ := json.Marshal(payload)
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(code)
	w.Write(response)
}

func respondWithError(w http.ResponseWriter, code int, msg string) {
	respondWithJson(w, code, map[string]string{"error": msg})
}

func main() {
	http.HandleFunc("/add-data/", addData)
	http.HandleFunc("/delete-data", deleteData)
	http.HandleFunc("/update-data/", updateData)
	http.HandleFunc("/search-by-city-category", searchByCity)
	http.HandleFunc("/search", search)
	http.HandleFunc("/add-data-category", addDataInCategory)
	http.HandleFunc("/delete-data-category", deleteDataInCategory)
	http.HandleFunc("/search-by-category", searchByCategory)
	http.HandleFunc("/update-data-category/", updateDataInCategory)
	//	http.HandleFunc("/pdf", writePdf)

	fmt.Println("Service Started at 8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}

/*func writePdf(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	if r.Method != "GET" {
		respondWithError(w, http.StatusBadRequest, "Invalid method")
		return
	}

	res, err := mongoDetails.WriteIntoPDF2()
	if err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	} else {
		respondWithJson(w, http.StatusAccepted, res)
	}

}*/
