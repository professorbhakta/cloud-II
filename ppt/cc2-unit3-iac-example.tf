terraform {
  required_version = ">= 1.5.0"

  required_providers {
    google = {
      source  = "hashicorp/google"
      version = "~> 5.0"
    }
  }
}

provider "google" {
  project = var.project_id
  region  = var.region
  zone    = var.zone
}

variable "project_id" {
  type = string
}

variable "region" {
  type    = string
  default = "asia-south1"
}

variable "zone" {
  type    = string
  default = "asia-south1-a"
}

variable "machine_type" {
  type    = string
  default = "e2-micro"
}

variable "app_port" {
  type    = number
  default = 8080
}

variable "labels" {
  type = map(string)
  default = {
    course      = "cc-ii"
    unit        = "unit-3"
    environment = "lab"
    managed_by  = "terraform"
  }
}

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
    access_config {}
  }

  metadata_startup_script = <<-EOT
    #!/bin/bash
    set -euo pipefail
    apt-get update -y
    apt-get install -y nginx
    echo "<h1>CC-II Unit III — IaC demo</h1>" > /var/www/html/index.html
    systemctl enable nginx
    systemctl restart nginx
  EOT

  allow_stopping_for_update = true
}

resource "google_compute_firewall" "fw_frontend" {
  name    = "cc2-fw-frontend"
  network = "default"

  allow {
    protocol = "tcp"
    ports    = [tostring(var.app_port), "80"]
  }

  target_tags   = ["frontend"]
  source_ranges = ["0.0.0.0/0"]
}

resource "google_storage_bucket" "artifacts" {
  name                        = "${var.project_id}-cc2-iac-artifacts"
  location                    = var.region
  force_destroy               = true
  uniform_bucket_level_access = true
  labels                      = var.labels

  versioning {
    enabled = true
  }
}

output "frontend_instance_name" {
  value = google_compute_instance.frontend.name
}

output "frontend_external_ip" {
  value = google_compute_instance.frontend.network_interface[0].access_config[0].nat_ip
}

output "firewall_name" {
  value = google_compute_firewall.fw_frontend.name
}

output "artifacts_bucket" {
  value = google_storage_bucket.artifacts.url
}
