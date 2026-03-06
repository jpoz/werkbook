package ooxml

import (
	"archive/zip"
	"fmt"
	"io"
)

const dynamicArrayMetadataXML = `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"></metadataType></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"></xda:dynamicArrayProperties></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"></rc></bk></cellMetadata></metadata>`

func writeDynamicArrayMetadata(zw *zip.Writer) error {
	return writeRawXML(zw, "xl/metadata.xml", dynamicArrayMetadataXML)
}

func writeRawXML(zw *zip.Writer, name, body string) error {
	w, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("create %s: %w", name, err)
	}
	if _, err := io.WriteString(w, xmlHeader); err != nil {
		return err
	}
	_, err = io.WriteString(w, body)
	return err
}
