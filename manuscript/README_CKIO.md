# Christian Knowledge Integration Ontology proof of concept

This package accompanies the manuscript **An AI-Assisted, Provenance-Aware Ontology for Christian Knowledge Integration**.

## Contents

- `Christian_Knowledge_Integration_Ontology_Manuscript.docx`: research manuscript.
- `ckio_prototype.ttl`: OWL/RDF ontology serialized as Turtle.
- `gospel_passage_dataset.csv`: 68 passage units assigned to 20 curated episode groupings.
- `evaluation_results.json`: deterministic integrity and competency-query outputs.
- `validate_ckio.py`: RDF syntax and assertion-completeness verification.
- `build_paper.py`: reproducible dataset, ontology, evaluation, and manuscript builder.

## Important scope limits

The episode labels are analytical groupings. Some group contextually related material rather than strict scene-by-scene narrative parallels. Passage ranges may overlap and therefore do not represent counts of distinct verses. The prototype has not received independent biblical-scholar or denominational validation and must not be treated as theological authority.

## Validation

The reported verification used Python 3.12.13 and RDFLib 7.1.4. From the project directory, run:

```bash
python validate_ckio.py
```

Expected core results are 780 explicit triples, 68 passages, 20 episode groupings, 68 computational assertions, zero incomplete assertions, and a `pass` status.

## License

The generated ontology, dataset, code, and documentation are provided under the Creative Commons Attribution 4.0 International license (CC BY 4.0). Biblical verse text is not included.
