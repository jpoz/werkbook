# ===========================================================================
# Werkbook Makefile — Build lifecycle entry point
# ===========================================================================
#
# Standard lifecycle:
#   test → lint → bench
#

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

BLUE   := \033[0;34m
GREEN  := \033[0;32m
YELLOW := \033[0;33m
RED    := \033[0;31m
NC     := \033[0m
BOLD   := \033[1m

# ---------------------------------------------------------------------------
# Print helpers
# ---------------------------------------------------------------------------

define print_stage
	@echo "$(BLUE)$(BOLD)▶ $(1)$(NC)"
endef

define print_success
	@echo "$(GREEN)✓ $(1)$(NC)"
endef

define print_info
	@echo "$(YELLOW)ℹ $(1)$(NC)"
endef

# ---------------------------------------------------------------------------
# Help target parser
# ---------------------------------------------------------------------------

define print_targets
	@awk \
		-v green="$(GREEN)" \
		-v bold="$(BOLD)" \
		-v nc="$(NC)" \
		'BEGIN {FS = ":.*?## "} { \
			if (/^[a-zA-Z_-]+:.*?##.*$$/) {printf "  %s%-20s%s %s\n", green, $$1, nc, $$2} \
			else if (/^## .*$$/) {printf "\n%s%s%s\n", bold, substr($$1,4), nc} \
		}' $(MAKEFILE_LIST)
	@echo ''
endef

.PHONY: all
all: help

# ============================================================================
## Help
# ============================================================================

.PHONY: help
help: ## Show this help
	@echo ""
	@echo "$(BOLD)Werkbook$(NC) — Go library for reading and writing Excel XLSX files"
	$(call print_targets)

# ============================================================================
## Setup
# ============================================================================

.PHONY: setup
setup: ## Install dependencies
	$(call print_stage,Downloading Go dependencies)
	go mod download
	$(call print_success,Setup complete!)

.PHONY: deps
deps: ## Download and tidy Go dependencies
	go mod download
	go mod tidy

# ============================================================================
## Testing
# ============================================================================

.PHONY: test
test: ## Run all unit tests
	$(call print_stage,Running tests)
	go test -v ./...

.PHONY: test-integration
test-integration: ## Run integration tests (requires LibreOffice)
	$(call print_stage,Running integration tests)
	go test -tags=integration -v ./...

.PHONY: bench
bench: ## Run benchmarks
	$(call print_stage,Running benchmarks)
	go test -bench=. -benchmem ./...

.PHONY: cover
cover: ## Run tests with coverage report
	$(call print_stage,Running tests with coverage)
	go test -coverprofile=coverage.out ./...
	go tool cover -html=coverage.out -o coverage.html
	$(call print_success,Coverage report: coverage.html)

# ============================================================================
## Linting & Formatting
# ============================================================================

.PHONY: lint
lint: ## Format and lint
	$(call print_stage,Formatting and linting)
	gofumpt -l -w .
	golangci-lint run --fix ./...

.PHONY: lint-check
lint-check: ## Run linters without fixing (for CI)
	golangci-lint run ./...

# ============================================================================
## Build
# ============================================================================

.PHONY: build
build: ## Verify the package compiles
	$(call print_stage,Building)
	go build ./...

.PHONY: clean
clean: ## Remove generated artifacts
	$(call print_stage,Cleaning)
	rm -f coverage.out coverage.html
