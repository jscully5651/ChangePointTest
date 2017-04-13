package main


import (
	"os"
	"log"
	"fmt"
	"time"
	"github.com/xlsx"
	"strconv"
)
const (
	filename = "changepoint.csv"
	weekhours = 144
)

type changepoint  struct {
	engagement	string
	project 	string
	task 		string
	company		string
	engineer	string
	billable	bool
	utilized	bool
	regHours	float64
	otHours		float64
	description	string
	date 		time.Time
}

type totalpercompany struct {
	engineer	string
	company		string
	engagement	string
	project 	string
	task 		string
	hoursPerClient	float64
}

type whitespace struct {
	project 	string
	hours		float64
	date		time.Time
}
// hours duration for a week is 144 since there is no time values and time starts at 0000.  Work week would be 40
type daterange struct{
	minDate		time.Time
	maxDate		time.Time
	expHours	float64
	numWeeks	float64
}

func (dr *daterange) lengthTime (){
	duration := dr.maxDate.Sub(dr.minDate).Hours()
	fmt.Println(duration)
	dr.numWeeks = duration/weekhours
	dr.expHours = 40 * dr.numWeeks
	fmt.Printf("number of weeks %.2f \t epected hours %.2f \n", dr.numWeeks, dr.expHours)


}

func (dr *daterange) setDateRange(cDate time.Time) {
	if cDate.After(dr.maxDate){
		dr.maxDate = cDate
		//fmt.Println("Am in max: ", dr.maxDate)

	}else if cDate.Before(dr.minDate) || dr.minDate.IsZero() {
		dr.minDate = cDate
		//fmt.Println("Am in min: ", dr.minDate)

	}

}

func checkProject ( engineer string, cpd changepoint, thr map[string][]totalpercompany){
	for idx, slice := range thr[engineer]{
		if slice.project == cpd.project && slice.project != "N/A" {
			if slice.engagement == cpd.engagement && slice.engagement != "N/A" {
				if slice.task == cpd.task {
					thr[engineer][idx].hoursPerClient += cpd.regHours + cpd.otHours
					return
				}
			}
		}else if slice.project == cpd.project && slice.project == "N/A" {
			if slice.engagement == cpd.engagement && slice.engagement == "N/A" {
				if slice.task == cpd.task {
					thr[engineer][idx].hoursPerClient += cpd.regHours
					return
				}
			}
		}

	}

	hours := cpd.regHours + cpd.otHours
	thr[engineer] = append(thr[engineer], totalpercompany{cpd.engineer,
	cpd.company,
	cpd.engagement,
	cpd.project,
	cpd.task,
	hours})
	return
}

func (dr *daterange) writeSheet ( mapChange map[string][]changepoint, mapTotal map[string][]totalpercompany) {
	var totalHours float64
	var utilHours  float64
	var billHours float64
	for _, tpc := range mapTotal {
		for _, slice := range tpc{
			slice.company = slice.company
		}
	}
	for eng, mcp := range mapChange{
		totalHours = 0.0
		billHours = 0.0
		utilHours = 0.0
		for _, cpSlice := range mcp{
			totalHours += cpSlice.otHours + cpSlice.regHours
			if cpSlice.billable {
				billHours += cpSlice.regHours + cpSlice.otHours
			}
			if cpSlice.utilized{
				utilHours += cpSlice.otHours + cpSlice.regHours
			}
		}

		fmt.Printf("%s \t\tutilized: \t %.2f \t billable: \t %2.f \n", eng, utilHours/(dr.numWeeks * dr.expHours)*100, billHours/(dr.numWeeks * dr.expHours)*100)
	}
}



func main () {
	drange := &daterange{}
	var mapCP = make(map[string][]changepoint, 25)
	var mapProject = make(map[string][] totalpercompany, 25)
	malformed := false
	excelFileName := "foo.xlsx"

	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
		os.Exit(1)
	}


	for _, sheet := range xlFile.Sheets {
		for ridx, row := range sheet.Rows {
			var resource string
			var cp changepoint
			if ridx == 0 {
				continue
			}
			malformed = false
			for cidx, cell := range row.Cells {
				text,_  := cell.String()
				if cidx ==7 && text == "" {
					log.Println("malformed row, please investigate row: ", ridx)
					malformed = true
					continue
				}
				if cidx == 7 {
					resource = text
					cp.engineer = text
				}
				if cidx == 0 {
					cp.company = text
				}
				if cidx == 1 {
					cp.engagement = text
				}
				if cidx == 2 {
					cp.project = text
				}
				if cidx == 3 {
					cp.task = text
				}
				if cidx == 8 {
					cp.date,err = cell.GetTime(false)
					if err != nil {
						log.Println("malformed time, please investigate row: ", ridx)
						malformed = true
						continue
					}
					drange.setDateRange(cp.date)
				}
				if cidx == 20 {
					if text == "Yes" {
						cp.billable = true
					}else {
						cp.billable = false
					}

				}
				if cidx == 22 {
					if text == "Yes" {
						cp.utilized = true
					}else {
						cp.utilized = false
					}
				}
				if cidx == 23 {
					cp.regHours, err = strconv.ParseFloat(text,64)
					if err != nil {
						log.Println(" regular hours did not parse, please investigate row: ", ridx)
						malformed = true
						continue
					}
				}
				if cidx == 24 {
					cp.otHours, err = strconv.ParseFloat(text,64)
					if err != nil {
						log.Println(" OT hours did not parse, please investigate row: ", ridx)
						malformed = true
						continue
					}
				}
				if cidx == 28{
					cp.description = text
				}
			}
			if !malformed {
				mapCP[resource] = append(mapCP[resource], cp)
				continue
			}
			checkProject (resource, cp, mapProject)
		}
	}
	drange.lengthTime()
	drange.writeSheet(mapCP, mapProject)
	//excelWSRName := "wsr.xlsx"
	//wsrFile, err := xlsx.OpenFile(excelWSRName)
	//if err != nil {
	//	log.Fatal(err)
	//	os.Exit(1)
	//}

}


