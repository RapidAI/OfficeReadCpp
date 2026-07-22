package main

import "testing"

func TestTokenOverlap(t *testing.T) {
	matched, reference, candidate, missing, extra := tokenOverlap("Alpha beta beta 123", "beta alpha gamma beta")
	if matched != 3 || reference != 4 || candidate != 4 {
		t.Fatalf("tokenOverlap() = (%d, %d, %d), want (3, 4, 4)", matched, reference, candidate)
	}
	if len(missing) != 1 || missing[0] != "123" || len(extra) != 1 || extra[0] != "gamma" {
		t.Fatalf("diagnostics = missing=%v extra=%v, want [123] and [gamma]", missing, extra)
	}
	recall, precision, f1 := scores(matched, reference, candidate)
	if recall != 0.75 || precision != 0.75 || f1 != 0.75 {
		t.Fatalf("scores() = (%v, %v, %v), want 0.75 each", recall, precision, f1)
	}
}

func TestScoresEmptyText(t *testing.T) {
	recall, precision, f1 := scores(0, 0, 0)
	if recall != 1 || precision != 1 || f1 != 1 {
		t.Fatalf("scores(0, 0, 0) = (%v, %v, %v), want (1, 1, 1)", recall, precision, f1)
	}
}
