# cloud-II — Course Materials (`cc2Material`)

All Cloud Computing – II (CC-II) teaching materials for FIT & CS, Parul University.

## Branches

| Branch | Purpose |
|--------|---------|
| `main` | Default branch |
| `cc2Material` | Consolidated course materials (this branch) |

## Folder layout

```
syllabus/       Official syllabus (PDF + units + practicals)
notes/unit-1/   Unit 1 study notes
notes/unit-2/   Unit 2 Disaster Recovery notes + GCP lab reference
notes/unit-3/   Unit 3 Infrastructure as Code (IaC) notes
notes/unit-4/   Unit 4 Kubernetes & Container Orchestration notes
notes/unit-5/   Unit 5 Introduction to CI/CD notes
notes/unit-6/   Unit 6 Cloud Application Integration notes
practicals/     Lab manuals for syllabus Practicals 1–11
ppt/            PowerPoint decks + Unit 4 visual assets
assignments/    Mid-term / assignment question papers
AGENTS.md       Guidance for Cursor / Cloud Agents
```

## Syllabus

- `syllabus/Cloud_Computing_II_Syllabus.pdf` — official PDF
- `syllabus/CC-II-Units.txt` — Units 1–6 (contents, weightage, hours)
- `syllabus/CC-II-Practicals.txt` — Practicals 1–11

## Notes

- `notes/unit-1/CC-II-Unit-I-Study-Notes.txt` — Unit 1 study notes
- `notes/unit-2/CC-II-Unit-II-Disaster-Recovery-in-Cloud.txt` — **Unit 2 syllabus notes (DR)**
- `notes/unit-2/CC-II-Unit-II-Networking-in-the-Cloud.txt` — extended GCP lab/reference guide
- `notes/unit-3/CC-II-Unit-III-Infrastructure-as-Code.txt` — **Unit 3 syllabus notes (IaC)**
- `notes/unit-4/CC-II-Unit-IV-Kubernetes-Container-Orchestration.txt` — **Unit 4 syllabus notes (Kubernetes)**
- `notes/unit-5/CC-II-Unit-V-Introduction-to-CI-CD.txt` — **Unit 5 syllabus notes (CI/CD)**
- `notes/unit-6/CC-II-Unit-VI-Cloud-Application-Integration.txt` — **Unit 6 syllabus notes (Integration)**

## Presentations

- `ppt/CC-II-Unit-II-Disaster-Recovery-in-Cloud.pptx` — Unit 2 expanded deck (~58 slides)
- `ppt/CC-II-Unit-III-Infrastructure-as-Code.pptx` — Unit 3 expanded deck (~51 slides)
- `ppt/cc2-unit3-iac-example.tf` — Unit 3 Terraform example (VM + firewall + bucket)
- `ppt/CC-II-Unit-IV-Kubernetes-Container-Orchestration.pptx` — Unit 4 expanded deck (~59 slides, with concept diagrams + GIFs)
- `ppt/unit4-assets/` — architecture, deploy, Service, Ingress, HPA, rolling-update visuals
- `ppt/CC-II-Unit-V-Introduction-to-CI-CD.pptx` — Unit 5 lecture deck (~37 slides, expanded teaching notes)

### Unit 2 deck topics

1. Disaster Recovery in Cloud  
2. Concepts of Backup & Disaster Recovery  
3. High Availability & Fault Tolerance in Cloud  
4. Disaster Recovery Strategies (RPO & RTO)  
5. Cloud-Based Backup Solutions & Snapshots  
6. Case Studies on Disaster Recovery Implementation  

### Unit 3 deck topics

1. Overview of Infrastructure as Code (IaC)  
2. Tools for IaC — Terraform, AWS CloudFormation, Ansible  
3. Configuration Management & Automation  
4. Writing & Managing Infrastructure as Code  
5. Best Practices for Versioning & Scaling Infrastructure  
6. Real-life case studies + exam revision  

### Unit 4 deck topics

1. Introduction to Kubernetes & its Architecture  
2. Setting Up & Managing Kubernetes Clusters  
3. Deploying Applications on Kubernetes  
4. Kubernetes Networking & Load Balancing  
5. Monitoring & Scaling Kubernetes Workloads  
6. Real-life case studies + exam revision  

### Unit 5 deck topics

1. Understanding DevOps and CI/CD Concepts  
2. CI/CD Pipelines: Overview & Implementation  
3. Tools — Jenkins, GitHub Actions (+ Copilot), GitLab CI  
4. Automated Testing & Deployment Strategies  
5. Best Practices for Secure & Scalable CI/CD Pipelines  
6. Real-life case studies + exam revision  

## Practicals

Lab manuals for all 11 official syllabus practicals:

- `practicals/CC-II-Practicals-Index.txt` — index, unit mapping, common setup
- `practicals/P01-GCP-Account-and-Linux-VM.txt`
- `practicals/P02-Docker-on-Compute-Engine.txt`
- `practicals/P03-Cloud-Run-Container-Deploy.txt`
- `practicals/P04-GKE-Deploy-and-Scale.txt`
- `practicals/P05-Cloud-Storage-Backup-Snapshots.txt`
- `practicals/P06-Cloud-Functions-Serverless.txt`
- `practicals/P07-Terraform-IaC.txt`
- `practicals/P08-Cloud-Build-CICD.txt`
- `practicals/P09-API-Gateway-CloudRun-PubSub.txt`
- `practicals/P10-GKE-Monitoring-Logging.txt`
- `practicals/P11-Disaster-Recovery-RPO-RTO.txt`

Each practical includes objectives, commands, verification checklist, deliverables, viva questions, and cleanup.

## Assignments

- `assignments/MiD1Assignment.txt` — MiD1 Assignment (Units 1–3 + Unit 4 starting topics; 98 marks)
- `assignments/Assignment-Solution.txt` — Model answer book for all MiD1 questions (Q.1–Q.14)
