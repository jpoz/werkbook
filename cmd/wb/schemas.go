package main

import "encoding/json"

// inputSchemas returns the JSON Schemas for the structured inputs that wb
// commands accept on stdin or via --spec / --patch / --config flags. The
// schemas are embedded in `wb capabilities` so agents can fetch the full
// input contract in one call.
//
// These are hand-written. They mirror the Go structs in patch.go,
// cmd_create.go, and cmd_check.go — keep them in sync when adding fields.
func inputSchemas() map[string]json.RawMessage {
	return map[string]json.RawMessage{
		"create_spec":  json.RawMessage(createSpecSchema),
		"patch_op":     json.RawMessage(patchOpSchema),
		"patch_array":  json.RawMessage(patchArraySchema),
		"check_config": json.RawMessage(checkConfigSchema),
	}
}

const patchOpSchema = `{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "patch_op",
  "description": "One operation: set a value/formula/style on a cell, set column width or row height, add or delete a sheet, or clear a cell or range. Unknown fields are rejected.",
  "type": "object",
  "additionalProperties": false,
  "properties": {
    "cell":         {"type": "string", "description": "A1-style cell reference, or a column letter for column_width, or a range like A1:B3 for clear/style."},
    "row":          {"type": "integer", "minimum": 1, "description": "1-based row index. Used with row_height."},
    "sheet":        {"type": "string", "description": "Target sheet name. Defaults to the first sheet (or the --sheet flag for edit)."},
    "type":         {"type": "string", "enum": ["date", "datetime", "time"], "description": "Interpret value as a temporal type. Value must be a JSON string. A default number-format style is applied unless 'style' is also supplied."},
    "value":        {"description": "Cell value: string, number, boolean, or null (clears). Combine with 'type' for typed dates/times."},
    "formula":      {"type": "string", "description": "Excel formula text. Leading '=' is tolerated and stripped."},
    "style":        {"type": "object", "description": "Style object. See readback shape under 'wb read --include-styles'."},
    "column_width": {"type": "number", "description": "Column width in Excel character units. 'cell' should be a column letter (e.g. 'B')."},
    "row_height":   {"type": "number", "description": "Row height in points. Requires 'row' to be set."},
    "add_sheet":    {"type": "string", "description": "Add a new sheet with this name."},
    "delete_sheet": {"type": "string", "description": "Delete the named sheet."},
    "clear":        {"type": "boolean", "description": "Clear the cell or range named by 'cell'."}
  }
}`

const patchArraySchema = `{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "patch_array",
  "description": "Input shape for 'wb edit': a JSON array of patch operations applied in order. With --atomic (the default) the file is saved only if every operation succeeds.",
  "type": "array",
  "items": {"$ref": "#/$defs/patch_op"},
  "$defs": {
    "patch_op": {"$ref": "patch_op"}
  }
}`

const createSpecSchema = `{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "create_spec",
  "description": "Input shape for 'wb create': declares sheets and seeds them with cell ops and/or row-oriented data blocks. Apply order is fixed: every op in 'cells' runs first, then 'rows' (in declaration order within each). When two ops target the same cell, the later op wins; 'wb create' surfaces these collisions in meta.warnings so an accidental clobber is visible. Unknown fields are rejected.",
  "type": "object",
  "additionalProperties": false,
  "properties": {
    "sheets": {
      "type": "array",
      "items": {"type": "string"},
      "description": "Sheet names in order. The first becomes the default sheet for cell ops without a 'sheet' field."
    },
    "cells": {
      "type": "array",
      "items": {"$ref": "patch_op"},
      "description": "Per-cell operations applied first. Same shape as the patch_op schema used by 'wb edit'."
    },
    "rows": {
      "type": "array",
      "items": {"$ref": "#/$defs/create_row"},
      "description": "Row-oriented data blocks applied after 'cells'. Each block lays values left-to-right, top-to-bottom from 'start'."
    }
  },
  "$defs": {
    "create_row": {
      "type": "object",
      "additionalProperties": false,
      "properties": {
        "sheet": {"type": "string", "description": "Target sheet. Defaults to the first sheet."},
        "start": {"type": "string", "description": "A1-style starting cell. Defaults to A1."},
        "data":  {
          "type": "array",
          "items": {"type": "array"},
          "description": "2D array of cell entries. Each entry is either a scalar (string/number/bool, or null which clears the target cell) or a JSON object with the per-cell fields {type, value, formula, style} from patch_op (cell/sheet are derived from the row geometry)."
        }
      },
      "required": ["data"]
    }
  }
}`

const checkConfigSchema = `{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "check_config",
  "description": "Optional config file for 'wb check --config'. Lets you set a global tolerance and ignore specific formulas or files.",
  "type": "object",
  "additionalProperties": false,
  "properties": {
    "tolerance": {
      "type": "number",
      "minimum": 0,
      "description": "Numeric tolerance for floating-point comparison. Cell mismatches within tolerance are treated as matches."
    },
    "ignore_formulas": {
      "type": "array",
      "items": {"type": "string"},
      "description": "Function-name patterns whose evaluations are skipped during comparison."
    },
    "ignore_files": {
      "type": "array",
      "items": {"type": "string"},
      "description": "Path-glob patterns. Matching files are skipped entirely."
    }
  }
}`
