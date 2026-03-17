package ooxml

import (
	"archive/zip"
	"fmt"
)

func writeRawXML(zw *zip.Writer, name, body string) error {
	return writeRawFile(zw, name, []byte(xmlHeader+body))
}

func writeRawFile(zw *zip.Writer, name string, body []byte) error {
	w, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("create %s: %w", name, err)
	}
	_, err = w.Write(body)
	return err
}
