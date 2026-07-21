# =============================================================================
# CC-II Unit III — Example Infrastructure as Code (Terraform)
# -----------------------------------------------------------------------------
# Purpose : Classroom / exam demo of declarative IaC on Google Cloud
# Course  : Cloud Computing – II (FIT & CS, Parul University)
# Lab link: Practical 7 — Automating Infrastructure Deployment with Terraform
#
# What this file declares (desired state):
#   1. A Compute Engine VM (frontend web server shape)
#   2. A firewall rule allowing TCP 8080 to VMs tagged "frontend"
#   3. A Cloud Storage bucket for application artifacts
#
# Teaching points:
#   • Declarative — we describe WHAT we want, not click-by-click HOW
#   • Variables parameterise project / region / size (Dev vs Prod)
#   • terraform plan  → preview creates / changes / destroys
#   • terraform apply → make the cloud match this file
#
# How to use (after gcloud auth and APIs enabled):
#   terraform init
#   terraform plan  -var="project_id=YOUR_GCP_PROJECT"
#   terraform apply -var="project_id=YOUR_GCP_PROJECT"
#   terraform destroy -var="project_id=YOUR_GCP_PROJECT"   # clean up lab
#
# Pair with: notes/unit-3/CC-II-Unit-III-Infrastructure-as-Code.txt
#            ppt/CC-II-Unit-III-Infrastructure-as-Code.pptx
# =============================================================================

terraform {
  required_version = ">= 1.5.0"

  required_providers {
    google = {
      source  = "hashicorp/google"
      version = "~> 5.0"
    }
  }

  # Classroom note:
  # For team / prod use a remote backend (e.g. GCS) + state locking.
  # Local state is OK for a personal lab demo only.
}

# -----------------------------------------------------------------------------
# Provider — which cloud & project Terraform will talk to
# -----------------------------------------------------------------------------
provider "google" {
  project = var.project_id
  region  = var.region
  zone    = var.zone
}

# -----------------------------------------------------------------------------
# Variables — same code shape; different inputs for Dev / Staging / Prod
# -----------------------------------------------------------------------------
variable "project_id" {
  description = "GCP project ID where resources will be created"
  type        = string
}

variable "region" {
  description = "GCP region"
  type        = string
  default     = "asia-south1"
}

variable "zone" {
  description = "GCP zone"
  type        = string
  default     = "asia-south1-a"
}

variable "machine_type" {
  description = "VM size (use e2-micro for free-tier-friendly labs)"
  type        = string
  default     = "e2-micro"
}

variable "app_port" {
  description = "Application port opened by the firewall"
  type        = number
  default     = 8080
}

variable "labels" {
  description = "Cost / ownership labels (FinOps + audit)"
  type        = map(string)
  default = {
    course      = "cc-ii"
    unit        = "unit-3"
    environment = "lab"
    managed_by  = "terraform"
  }
}

# -----------------------------------------------------------------------------
# Resource 1 — Compute Engine VM (provisioning)
# Conceptual parallel: gcloud compute instances create …
# -----------------------------------------------------------------------------
resource "google_compute_instance" "frontend" {
  name         = "cc2-frontend-1"
  machine_type = var.machine_type
  zone         = var.zone
  tags         = ["frontend"]
  labels       = var.labels

  boot_disk {
    initialize_params {
      image = "debian-cloud/debian-12"
      size  = 10
      type  = "pd-balanced"
    }
  }

  network_interface {
    network = "default"

    # Ephemeral public IP — fine for a lab demo.
    # Production often prefers private IP + IAP / load balancer only.
    access_config {}
  }

  # Startup script = configuration management pattern (Unit III Topic 3).
  # In larger projects this is often replaced by Ansible or a golden image.
  metadata_startup_script = <<-EOT
    #!/bin/bash
    set -euo pipefail
    apt-get update -y
    apt-get install -y nginx
    echo "<h1>CC-II Unit III — IaC demo (Terraform)</h1>" > /var/www/html/index.html
    systemctl enable nginx
    systemctl restart nginx
  EOT

  # Classroom tip: allow_stopping_for_update lets Terraform change machine_type
  # more easily during labs (still read the plan carefully).
  allow_stopping_for_update = true
}

# -----------------------------------------------------------------------------
# Resource 2 — Firewall rule (provisioning)
# Conceptual parallel: Fancy Store fw-fe style rule in Unit III notes / PPT
# -----------------------------------------------------------------------------
resource "google_compute_firewall" "fw_frontend" {
  name    = "cc2-fw-frontend"
  network = "default"

  description = "Allow app traffic to VMs tagged frontend (lab demo)"

  allow {
    protocol = "tcp"
    ports    = [tostring(var.app_port), "80"]
  }

  target_tags   = ["frontend"]
  source_ranges = ["0.0.0.0/0"] # DEMO ONLY — tighten for real production!

  # Exam talking point:
  # Peer review should catch overly open source_ranges before apply.
}

# -----------------------------------------------------------------------------
# Resource 3 — Cloud Storage bucket for artifacts
# Conceptual parallel: store app packages / backups of configs
# -----------------------------------------------------------------------------
resource "google_storage_bucket" "artifacts" {
  name                        = "${var.project_id}-cc2-iac-artifacts"
  location                    = var.region
  force_destroy               = true # lab convenience — avoid in real prod
  uniform_bucket_level_access = true

  labels = var.labels

  versioning {
    enabled = true
  }
}

# -----------------------------------------------------------------------------
# Outputs — useful values after terraform apply (Verify step)
# -----------------------------------------------------------------------------
output "frontend_instance_name" {
  description = "Name of the frontend VM"
  value       = google_compute_instance.frontend.name
}

output "frontend_external_ip" {
  description = "Public IP to open in a browser (lab)"
  value       = google_compute_instance.frontend.network_interface[0].access_config[0].nat_ip
}

output "firewall_name" {
  description = "Firewall rule created by this stack"
  value       = google_compute_firewall.fw_frontend.name
}

output "artifacts_bucket" {
  description = "GCS bucket for lab artifacts"
  value       = google_storage_bucket.artifacts.url
}

# =============================================================================
# Exam / viva — one-minute explanation of THIS file
# -----------------------------------------------------------------------------
# 1. This is declarative IaC: desired VM + firewall + bucket are declared.
# 2. terraform plan shows what will be created before any change.
# 3. terraform apply creates the resources via Google APIs (provisioning).
# 4. The startup script configures software on the VM (CM pattern).
# 5. Variables let the same file serve Dev (e2-micro) or a larger lab size.
# 6. Git should store this file; never commit service-account keys / passwords.
# =============================================================================
