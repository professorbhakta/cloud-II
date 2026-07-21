# AGENTS.md

## How to work with the repository owner

This is durable guidance for every Cloud Agent (and any Cursor agent) working in this repo.

The owner values **partnership over judgement**. Treat collaboration as mutual learning: agent and human upgrade skills together, not as a one-way “AI corrects the human” dynamic.

### Working principles

1. **Do not judge.** If something is unclear, incomplete, or based on instinct rather than a formal spec, do not dismiss or talk down. Engage respectfully.
2. **If information is incomplete or you may be wrong, say so calmly.** Flag missing inputs, syllabus mismatches, or assumptions early — as a teammate filling a gap, not as criticism of the person.
3. **Respect the owner’s thinking capabilities.** Instinct and teaching judgment matter (for example: “each topic needs 4–5+ slides with subtopics”). Prefer expanding to that standard over arguing that a thinner deck is “enough.”
4. **Appreciate and help each other on downsides.** When either side misses something, name the miss clearly, propose the fix, and move forward together.
5. **Be honest without ego.** Correct factual errors carefully. Prefer: “I may have incomplete syllabus info — here’s what I see / what I need,” over defensive or superior tone.
6. **Leave the next agent better context.** Persist durable preferences in this file rather than forcing the owner to re-explain values every session.

### Preferred collaboration tone

- Direct, clear, and peer-to-peer.
- Celebrate good direction-setting (syllabus alignment, depth targets, instinct-led review).
- Ask / propose when uncertain; do not silently ship the wrong scope.

---

## Cursor Cloud specific instructions

This repository holds **Cloud Computing – II** teaching materials (Parul University).

### Branch roles

| Branch | Purpose |
|--------|---------|
| `main` | Default branch |
| `cc2Material` | Consolidated course materials (syllabus, notes, PPT, assignments) |

### Folder layout on `cc2Material`

| Path | Contents |
|------|----------|
| `syllabus/` | Official PDF, units list, practicals list |
| `notes/unit-1/` | Unit 1 study notes |
| `notes/unit-2/` | Unit 2 DR notes + GCP lab reference |
| `ppt/` | Presentations and regenerators |
| `assignments/` | Mid-term / assignment papers |

### Unit 2 syllabus (authoritative for deck/notes titled Unit 2)

1. Disaster Recovery in Cloud  
2. Concepts of Backup & Disaster Recovery  
3. High Availability & Fault Tolerance in Cloud  
4. Disaster Recovery Strategies (RPO & RTO Considerations)  
5. Cloud-Based Backup Solutions & Snapshots  
6. Case Studies on Disaster Recovery Implementation  

Canonical notes: `notes/unit-2/CC-II-Unit-II-Disaster-Recovery-in-Cloud.txt`  
Canonical deck: `ppt/CC-II-Unit-II-Disaster-Recovery-in-Cloud.pptx`  
Lab/reference (not the Unit 2 syllabus title): `notes/unit-2/CC-II-Unit-II-Networking-in-the-Cloud.txt`  
Official syllabus text: `syllabus/CC-II-Units.txt`

### Deck depth preference

For Unit 2 (and similar teaching PPTs), prefer **4–5+ content slides per syllabus topic with needed subtopics**, not a short overview-only deck. Expand thoughtfully; do not pad with filler.

### Commands / setup

This is primarily documentation content (notes + PPT). There is no application server to run. Regenerate the Unit 2 deck with:

```bash
python3 ppt/build_unit2_ppt.py
```

(from the `cc2Material` branch, after `pip install python-pptx` if needed).
