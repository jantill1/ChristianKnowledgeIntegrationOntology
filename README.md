# Christian Knowledge Integration Ontology (CKIO)

CKIO is a provenance-aware OWL/RDF proof of concept for integrating Christian
knowledge without treating scriptural evidence, computational assertions, and
interpretive claims as equivalent facts.

The current release models the four canonical Gospels through 20 curated
analytical episode groupings, 68 passage units, 68 typed mappings, and 68
reified computational assertions. It is an engineering prototype, not a
theological authority or a claim of doctrinal neutrality.

## Repository contents

- `ontology/ckio_prototype.ttl` — ontology and instance graph in Turtle.
- `data/gospel_passage_dataset.csv` — structured passage-level dataset.
- `results/evaluation_results.json` — competency-query outcomes.
- `results/validation_report.json` — machine-readable validation report.
- `scripts/validate_ckio.py` — reproducible RDF and integrity checks.
- `scripts/build_paper.py` — manuscript generator.
- `manuscript/` — journal-submission manuscript and supporting documentation.

## Validate

Requires Python 3.12 or later and RDFLib 7.1.4.

```bash
python scripts/validate_ckio.py
```

The verified release parses as 780 explicit RDF triples and passes all declared
integrity checks and six fixture-based competency queries. These checks establish
implementation consistency; they do not constitute independent theological or
external-validity review.

## Citation and license

Citation metadata are provided in `CITATION.cff`. All original materials are
released under the Creative Commons Attribution 4.0 International license.

## Research status

Before relying on the passage mappings for scholarly conclusions, users should
seek review by qualified biblical scholars and document the interpretive and
denominational perspectives applied.
