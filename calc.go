package werkbook

import "github.com/jpoz/werkbook/ooxml"

// CalcProperties exposes the workbook's <calcPr> settings.
type CalcProperties struct {
	Mode           string
	ID             int
	FullCalcOnLoad bool
	ForceFullCalc  bool
	Completed      bool
}

// SetDate1904 switches the workbook date system used for date serialization and formula formatting.
func (f *File) SetDate1904(enabled bool) {
	if f.date1904 == enabled {
		return
	}
	f.date1904 = enabled
	f.rebuildFormulaState()
}

// CalcProperties returns the workbook calculation properties.
func (f *File) CalcProperties() CalcProperties {
	return f.calcProps
}

// SetCalcProperties replaces the workbook calculation properties.
func (f *File) SetCalcProperties(props CalcProperties) {
	f.calcProps = props
}

func calcPropsFromData(data ooxml.CalcPropertiesData) CalcProperties {
	return CalcProperties{
		Mode:           data.Mode,
		ID:             data.ID,
		FullCalcOnLoad: data.FullCalcOnLoad,
		ForceFullCalc:  data.ForceFullCalc,
		Completed:      data.Completed,
	}
}

func (p CalcProperties) toData() ooxml.CalcPropertiesData {
	return ooxml.CalcPropertiesData{
		Mode:           p.Mode,
		ID:             p.ID,
		FullCalcOnLoad: p.FullCalcOnLoad,
		ForceFullCalc:  p.ForceFullCalc,
		Completed:      p.Completed,
	}
}
