"""Independent syntax and structural validation for the CKIO prototype."""
from pathlib import Path
import json
from rdflib import Graph, Namespace, RDF

ROOT = Path(__file__).parent
OUT = ROOT / "output"
graph = Graph().parse(OUT / "ckio_prototype.ttl", format="turtle")
CKIO = Namespace("https://example.org/ckio/")

passages = set(graph.subjects(RDF.type, CKIO.Passage))
episodes = set(graph.subjects(RDF.type, CKIO.Episode))
assertions = set(graph.subjects(RDF.type, CKIO.ComputationalAssertion))

required = [CKIO.hasEvidence, CKIO.proposesEpisode, CKIO.usesPredicate, CKIO.hasStatus]
incomplete = [str(a) for a in assertions if any(not list(graph.objects(a, p)) for p in required)]

report = {
    "parser": "RDFLib",
    "rdflib_version": __import__("rdflib").__version__,
    "explicit_triples": len(graph),
    "passages": len(passages),
    "episodes": len(episodes),
    "computational_assertions": len(assertions),
    "incomplete_assertions": incomplete,
    "status": "pass" if not incomplete and len(passages) == 68 and len(episodes) == 20 and len(assertions) == 68 else "fail",
}
print(json.dumps(report, indent=2))
(OUT / "validation_report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
if report["status"] != "pass":
    raise SystemExit(1)
