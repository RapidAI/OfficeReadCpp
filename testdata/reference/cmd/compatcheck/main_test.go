package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestMarkdownImageReferenceCount(t *testing.T) {
	markdown := "plain ![first](images/one.png) text ![broken] and ![second](images/two.png)"
	if got, want := markdownImageReferenceCount(markdown), 2; got != want {
		t.Fatalf("markdownImageReferenceCount() = %d, want %d", got, want)
	}
}

func TestCheckFilesPreservesInputOrder(t *testing.T) {
	dir := t.TempDir()
	paths := []string{
		filepath.Join(dir, "first.txt"),
		filepath.Join(dir, "second.txt"),
	}
	for _, path := range paths {
		if err := os.WriteFile(path, []byte("not an Office file"), 0o600); err != nil {
			t.Fatal(err)
		}
	}
	results := checkFiles(paths, 1, false, 2)
	if len(results) != len(paths) {
		t.Fatalf("got %d results, want %d", len(results), len(paths))
	}
	for i, result := range results {
		if result.Path != paths[i] {
			t.Errorf("result %d path = %q, want %q", i, result.Path, paths[i])
		}
	}
}
