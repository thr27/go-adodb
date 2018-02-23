package main

//package main

import (
	"database/sql"
	"fmt"
	"os"

	ole "github.com/go-ole/go-ole"
	_ "github.com/go-ole/go-ole/oleutil"
	_ "github.com/thr27/go-adodb"
	//_ "github.com/mattn/go-adodb"
)

var (
	db         *sql.DB
	checkError = func(err error, num int) {
		if err != nil {
			fmt.Println("Error:")
			fmt.Println(err, num)
			os.Exit(1)
		}
	}
)

func main() {
	fmt.Println("Hallo ...")
	ole.CoInitialize(0)
	// Replace the DBQ value with the name of your ODBC data source.
	fmt.Println("Open DB ...")

	db, err := sql.Open("adodb", "Provider=Advantage.OLEDB.1;Data Source='D:\\proquest\\airquest\\db\\'; ServerType=ADS_LOCAL_SERVER; TableType=ADS_CDX; ")
	checkError(err, 1)

	fmt.Println("Query DB ...")

	rows, err := db.Query("SELECT count(*) as cnt FROM QUEUE WHERE CQUEUENO_ = ?", "ARCHIV")

	checkError(err, 2)

	for rows.Next() {

		var cnt int

		err = rows.Scan(&cnt)
		if err != nil {
			fmt.Println("scan", err)
			return
		}
		fmt.Println("Count:", cnt)

	}
	defer rows.Close()

	args := make([]interface{}, 2, 2)
	args[0] = "HALLO"
	args[1] = "FTI4092222035955639"

	s := fmt.Sprintf("UPDATE QUEUE set cupduser_ = ? WHERE csurrogat_ = ?")
	_, err = db.Exec(s, args...)

	defer db.Close()
}
