package formula

import (
	"errors"
	"testing"
)

func TestErrorValueFromErr(t *testing.T) {
	cases := []struct {
		name string
		err  error
		want ErrorValue
	}{
		{"nil_defaults_to_NAME", nil, ErrValNAME},
		{"plain_error_defaults_to_NAME", errors.New("parse: unexpected EOF"), ErrValNAME},
		{"formula_too_large_maps_to_VALUE", ErrFormulaTooLarge, ErrValVALUE},
		{
			"wrapped_eval_error_preserves_code_REF",
			WrapEvalError(ErrValREF, errors.New("circular reference")),
			ErrValREF,
		},
		{
			"wrapped_eval_error_preserves_code_VALUE",
			WrapEvalError(ErrValVALUE, errors.New("stack underflow")),
			ErrValVALUE,
		},
		{
			"nested_wrapped_error_still_classified",
			WrapEvalError(ErrValDIV0, errors.New("divide by zero")),
			ErrValDIV0,
		},
		{
			"formula_too_large_outranks_NAME_default",
			errors.Join(ErrFormulaTooLarge, errors.New("other")),
			ErrValVALUE,
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := ErrorValueFromErr(tc.err)
			if got != tc.want {
				t.Errorf("ErrorValueFromErr(%v) = %v (%s), want %v (%s)",
					tc.err, got, got, tc.want, tc.want)
			}
		})
	}
}

// TestWrapEvalErrorNilReturnsNil locks down the "safe to call in return
// statements" contract so existing early-return sites can swap err →
// WrapEvalError(code, err) without adding a guard for the nil case.
func TestWrapEvalErrorNilReturnsNil(t *testing.T) {
	if got := WrapEvalError(ErrValNAME, nil); got != nil {
		t.Errorf("WrapEvalError(code, nil) = %v, want nil", got)
	}
}
