"""
Experimentos computacionales: Búsqueda Tabú + ILS sobre keller4 y C125.9.

Genera:
  - results_summary.json / results_summary.csv: estadísticas agregadas por instancia.
  - convergencia_<instancia>.png: curva de convergencia (mejor f(S) vs iteración) de una corrida representativa.
  - grafo_clique_<instancia>_tabu.png: visualización del grafo con el clique encontrado resaltado.
  - comparacion_tiempos.png: comparación de tiempos metaheurística vs método exacto (Informe 1).
"""
import json
import statistics as st
import time

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np

from mcp_tabu_ils import load_dimacs, tabu_ils, verify_clique

N_RUNS = 30
MAX_ITER = 8000
TENURE = 10
STALL_LIMIT = 200

INSTANCES = [
    {"name": "keller4", "file": "keller4.clq", "known_optimum": 11,
     "exact_ref": "HiGHS (Informe 1)", "exact_time_s": 15.0, "exact_optimal": True},
    {"name": "C125.9", "file": "C125.9.clq", "known_optimum": 34,
     "exact_ref": "B&B manual (PyBnB)", "exact_time_s": 300.0, "exact_optimal": False},
]


def iters_time_to_best(history, best_size):
    for it, t, cur, best in history:
        if best == best_size:
            return it, t
    return history[-1][0], history[-1][1]


def run_instance(inst):
    """Ejecuta N_RUNS corridas independientes (semillas 0..N_RUNS-1).

    La corrida "representativa" usada para las figuras de convergencia y de
    clique (convergencia_<inst>.png, grafo_clique_<inst>_tabu.png) es,
    determinísticamente, la que tardó más iteraciones en alcanzar su mejor
    solución (argmax de iters_to_best) entre las N_RUNS corridas. Esta es la
    corrida con la trayectoria de convergencia más larga e ilustrativa (más
    plateaus y perturbaciones visibles antes de converger), a diferencia de
    n_restarts, que puede ser alto incluso en corridas que alcanzan el óptimo
    casi de inmediato y luego perturban sin efecto visible en el incumbente.
    En caso de empate se toma la semilla más baja.
    """
    n, A = load_dimacs(inst["file"])
    m = int(A.sum() // 2)
    density = m / (n * (n - 1) / 2)

    runs = []
    all_histories = {}
    all_best_S = {}
    for seed in range(N_RUNS):
        res = tabu_ils(A, n, max_iterations=MAX_ITER, tenure=TENURE,
                        stall_limit=STALL_LIMIT, seed=seed)
        assert verify_clique(A, res.best_S), "Clique inválido detectado"
        it_best, t_best = iters_time_to_best(res.history, res.best_size)
        runs.append({
            "seed": seed,
            "best_size": res.best_size,
            "iters_to_best": it_best,
            "time_to_best": t_best,
            "total_time": res.elapsed,
            "n_restarts": res.n_restarts,
            "hit_optimum": res.best_size == inst["known_optimum"],
        })
        all_histories[seed] = res.history
        all_best_S[seed] = res.best_S

    rep_seed = max(range(N_RUNS), key=lambda s: (runs[s]["iters_to_best"], -s))
    rep_history = all_histories[rep_seed]
    rep_best_S = all_best_S[rep_seed]
    print(f"  Semilla representativa ({inst['name']}): {rep_seed} "
          f"(iters_to_best={runs[rep_seed]['iters_to_best']}, "
          f"{runs[rep_seed]['n_restarts']} perturbaciones)")

    sizes = [r["best_size"] for r in runs]
    times_to_best = [r["time_to_best"] for r in runs]
    iters_to_best = [r["iters_to_best"] for r in runs]
    hits = [r["hit_optimum"] for r in runs]

    summary = {
        "instance": inst["name"],
        "n_vertices": n,
        "n_edges": m,
        "density": round(density, 4),
        "known_optimum": inst["known_optimum"],
        "n_runs": N_RUNS,
        "best_size_mean": round(st.mean(sizes), 3),
        "best_size_std": round(st.pstdev(sizes), 3),
        "best_size_max": max(sizes),
        "success_rate": round(sum(hits) / N_RUNS, 3),
        "time_to_best_mean_s": round(st.mean(times_to_best), 4),
        "time_to_best_std_s": round(st.pstdev(times_to_best), 4),
        "iters_to_best_mean": round(st.mean(iters_to_best), 1),
        "exact_ref": inst["exact_ref"],
        "exact_time_s": inst["exact_time_s"],
        "exact_optimal": inst["exact_optimal"],
        "representative_seed": rep_seed,
        "representative_n_restarts": runs[rep_seed]["n_restarts"],
        "representative_iters_to_best": runs[rep_seed]["iters_to_best"],
    }
    return summary, runs, rep_history, rep_best_S, n, A


def plot_convergence(history, inst_name, known_optimum, path):
    its = [h[0] for h in history]
    best = [h[3] for h in history]
    fig, ax = plt.subplots(figsize=(7, 4.5))
    ax.plot(its, best, color="#1f4e79", linewidth=1.8, label="Mejor $f(S)$ (incumbente)")
    ax.axhline(known_optimum, color="#c0392b", linestyle="--", linewidth=1.4,
               label=f"Óptimo conocido ($\\omega={known_optimum}$)")
    ax.set_xlabel("Iteración")
    ax.set_ylabel("Tamaño del clique")
    ax.set_title(f"Convergencia de Búsqueda Tabú + ILS — {inst_name}")
    ax.legend(loc="lower right")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def plot_clique_graph(A, S, inst_name, path):
    n = A.shape[0]
    G = nx.from_numpy_array(A)
    clique_nodes = set(np.where(S)[0])
    other_nodes = [v for v in G.nodes() if v not in clique_nodes]
    clique_edges = [(u, v) for u, v in G.edges() if u in clique_nodes and v in clique_nodes]
    other_edges = [(u, v) for u, v in G.edges() if not (u in clique_nodes and v in clique_nodes)]

    pos = nx.spring_layout(G, seed=42, k=0.3)
    fig, ax = plt.subplots(figsize=(9, 9))
    nx.draw_networkx_edges(G, pos, edgelist=other_edges, alpha=0.03, edge_color="gray", ax=ax)
    nx.draw_networkx_edges(G, pos, edgelist=clique_edges, alpha=0.7, edge_color="#c0392b", width=1.6, ax=ax)
    nx.draw_networkx_nodes(G, pos, nodelist=other_nodes, node_size=25, node_color="#a9cce3", ax=ax)
    nx.draw_networkx_nodes(G, pos, nodelist=list(clique_nodes), node_size=90, node_color="#c0392b", ax=ax)
    ax.set_title(f"Clique encontrado por Tabú+ILS — {inst_name} (|S|={len(clique_nodes)})")
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def plot_time_comparison(summaries, path):
    names = [s["instance"] for s in summaries]
    exact_times = [s["exact_time_s"] for s in summaries]
    meta_times = [s["time_to_best_mean_s"] for s in summaries]
    exact_labels = [("óptimo probado" if s["exact_optimal"] else "timeout, sin\nprobar óptimo") for s in summaries]

    x = np.arange(len(names))
    width = 0.35
    fig, ax = plt.subplots(figsize=(7, 4.8))
    ax.set_yscale("log")
    b1 = ax.bar(x - width/2, exact_times, width, label="Método exacto (Informe 1)", color="#2874a6")
    b2 = ax.bar(x + width/2, meta_times, width, label="Tabú + ILS (Informe 2)", color="#c0392b")
    for i, (bar, lab) in enumerate(zip(b1, exact_labels)):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height()*1.15, lab,
                ha="center", va="bottom", fontsize=8)
    for bar, s in zip(b2, summaries):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height()*1.15,
                f"{bar.get_height():.3f}s", ha="center", va="bottom", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(names)
    ax.set_ylabel("Tiempo hasta la mejor solución (s, escala log)")
    ax.set_title("Tiempo de cómputo: método exacto vs metaheurística", pad=14)
    ax.set_ylim(top=max(exact_times) * 8)  # espacio para que las etiquetas de las barras no choquen con el título
    ax.legend()
    ax.grid(True, axis="y", alpha=0.3, which="both")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def main():
    all_summaries = []
    for inst in INSTANCES:
        print(f"Procesando {inst['name']} ({N_RUNS} corridas x {MAX_ITER} iteraciones)...")
        t0 = time.time()
        summary, runs, rep_history, rep_best_S, n, A = run_instance(inst)
        print(f"  -> {time.time()-t0:.1f}s totales. Resumen: {summary}")
        all_summaries.append(summary)

        plot_convergence(rep_history, inst["name"], inst["known_optimum"],
                          f"convergencia_{inst['name']}.png")
        plot_clique_graph(A, rep_best_S, inst["name"],
                           f"grafo_clique_{inst['name']}_tabu.png")

        with open(f"runs_{inst['name']}.json", "w") as f:
            json.dump(runs, f, indent=2)

    with open("results_summary.json", "w") as f:
        json.dump(all_summaries, f, indent=2)

    plot_time_comparison(all_summaries, "comparacion_tiempos.png")

    print("\n=== RESUMEN FINAL ===")
    for s in all_summaries:
        print(s)


if __name__ == "__main__":
    main()
