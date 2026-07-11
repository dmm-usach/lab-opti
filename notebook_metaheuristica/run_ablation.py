"""
Ablación: Búsqueda Tabú pura (sin reinicio) vs. con reinicio tipo ILS.

Reproduce exactamente la Tabla 3 del Informe 2 (Sección 4.3): se repiten las
mismas 30 corridas (semillas 0-29) de `run_experiments.py` sobre keller4 y
C125.9, pero con `stall_limit` desactivado (`stall_limit -> infinito`), de
modo que el mecanismo de perturbación tipo ILS nunca se dispara y el
algoritmo se reduce a una Búsqueda Tabú pura. El resultado se contrasta
contra `results_summary.json` (que sí usa el reinicio).

Genera: ablation_sin_reinicio.json
"""
import json
import statistics as st

from mcp_tabu_ils import load_dimacs, tabu_ils, verify_clique

N_RUNS = 30
MAX_ITER = 8000
TENURE = 10
STALL_LIMIT_DISABLED = 10**9  # equivalente a stall_limit -> infinito

INSTANCES = [
    {"name": "keller4", "file": "keller4.clq", "known_optimum": 11},
    {"name": "C125.9", "file": "C125.9.clq", "known_optimum": 34},
]


def run_ablation_instance(inst):
    n, A = load_dimacs(inst["file"])
    sizes = []
    hits = 0
    for seed in range(N_RUNS):
        res = tabu_ils(A, n, max_iterations=MAX_ITER, tenure=TENURE,
                        stall_limit=STALL_LIMIT_DISABLED, seed=seed)
        assert verify_clique(A, res.best_S), "Clique inválido detectado"
        sizes.append(res.best_size)
        if res.best_size == inst["known_optimum"]:
            hits += 1
    return {
        "success_rate_sin_reinicio": round(hits / N_RUNS, 3),
        "best_size_mean_sin_reinicio": round(st.mean(sizes), 3),
        "best_size_std_sin_reinicio": round(st.pstdev(sizes), 3),
    }


def main():
    results = {}
    for inst in INSTANCES:
        print(f"Procesando {inst['name']} sin reinicio ({N_RUNS} corridas x {MAX_ITER} iteraciones)...")
        results[inst["name"]] = run_ablation_instance(inst)
        print(f"  -> {results[inst['name']]}")

    with open("ablation_sin_reinicio.json", "w") as f:
        json.dump(results, f, indent=2)

    print("\n=== RESUMEN ABLACIÓN (sin reinicio) ===")
    for name, r in results.items():
        print(name, r)


if __name__ == "__main__":
    main()
