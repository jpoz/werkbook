package werkbook

import (
	"time"

	"github.com/jpoz/werkbook/ooxml"
)

// CoreProperties exposes the workbook's supported OPC core properties.
// Zero Created or Modified times mean the property is absent.
type CoreProperties struct {
	Title          string
	Subject        string
	Creator        string
	Description    string
	Identifier     string
	Language       string
	Keywords       string
	Category       string
	ContentStatus  string
	Version        string
	Revision       string
	LastModifiedBy string
	Created        time.Time
	Modified       time.Time
}

// CoreProperties returns the workbook's OPC core properties.
func (f *File) CoreProperties() CoreProperties {
	return f.coreProps
}

// SetCoreProperties replaces the workbook's OPC core properties.
func (f *File) SetCoreProperties(props CoreProperties) {
	f.coreProps = props
	f.corePropsDirty = true
}

func corePropsFromData(data ooxml.CorePropertiesData) CoreProperties {
	return CoreProperties{
		Title:          data.Title,
		Subject:        data.Subject,
		Creator:        data.Creator,
		Description:    data.Description,
		Identifier:     data.Identifier,
		Language:       data.Language,
		Keywords:       data.Keywords,
		Category:       data.Category,
		ContentStatus:  data.ContentStatus,
		Version:        data.Version,
		Revision:       data.Revision,
		LastModifiedBy: data.LastModifiedBy,
		Created:        data.Created,
		Modified:       data.Modified,
	}
}

func (p CoreProperties) toData() ooxml.CorePropertiesData {
	return ooxml.CorePropertiesData{
		Title:          p.Title,
		Subject:        p.Subject,
		Creator:        p.Creator,
		Description:    p.Description,
		Identifier:     p.Identifier,
		Language:       p.Language,
		Keywords:       p.Keywords,
		Category:       p.Category,
		ContentStatus:  p.ContentStatus,
		Version:        p.Version,
		Revision:       p.Revision,
		LastModifiedBy: p.LastModifiedBy,
		Created:        p.Created,
		Modified:       p.Modified,
	}
}
