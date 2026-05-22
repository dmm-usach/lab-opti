"""
Comparativa HiGHS vs GLPK sobre la instancia keller4 del benchmark DIMACS.

Resuelve la formulación ILP del Maximum Clique Problem con ambos solvers
y reporta tiempo, status y omega encontrado. GLPK se corre con un límite
de tiempo de 300 s (5 min) para reflejar el escenario descrito en el informe.
"""
import time
import networkx as nx
import pyomo.environ as pyo

INSTANCE = "keller4.clq"
TIME_LIMIT_GLPK = 300  # segundos


def load_dimacs(path):
    G = nx.Graph()
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("c"):
                continue
            if line.startswith("p"):
                _, _, n, m = line.split()
                G.add_nodes_from(range(1, int(n) + 1))
            elif line.startswith("e"):
                _, u, v = line.split()
                G.add_edge(int(u), int(v))
    return G


def build_mcp_model(G):
    model = pyo.ConcreteModel("MaximumClique")
    model.V = pyo.Set(initialize=list(G.nodes()))
    model.x = pyo.Var(model.V, domain=pyo.Binary)
    model.obj = pyo.Objective(
        expr=sum(model.x[i] for i in model.V),
        sense=pyo.maximize,
    )
    model.nonadj = pyo.ConstraintList()
    nodes = list(G.nodes())
    for i_idx, i in enumerate(nodes):
        for j in nodes[i_idx + 1:]:
            if not G.has_edge(i, j):
                model.nonadj.add(model.x[i] + model.x[j] <= 1)
    return model


def solve_with(solver_name, model, **opts):
    solver = pyo.SolverFactory(solver_name)
    for k, v in opts.items():
        solver.options[k] = v
    t0 = time.time()
    res = solver.solve(model, tee=False)
    elapsed = time.time() - t0
    status = str(res.solver.termination_condition)
    try:
        obj = pyo.value(model.obj)
    except Exception:
        obj = None
    return elapsed, status, obj


def main():
    print(f"Cargando instancia {INSTANCE}...")
    G = load_dimacs(INSTANCE)
    print(f"  |V| = {G.number_of_nodes()}, |E| = {G.number_of_edges()}")

    print("\n=== HiGHS ===")
    model_h = build_mcp_model(G)
    print(f"  Variables: {len(model_h.V)}, Restricciones: {len(model_h.nonadj)}")
    t, st, obj = solve_with("appsi_highs", model_h)
    print(f"  Tiempo: {t:.2f} s | Status: {st} | Omega: {obj}")

    print(f"\n=== GLPK (time limit = {TIME_LIMIT_GLPK} s) ===")
    model_g = build_mcp_model(G)
    t, st, obj = solve_with("glpk", model_g, tmlim=TIME_LIMIT_GLPK)
    print(f"  Tiempo: {t:.2f} s | Status: {st} | Mejor encontrado: {obj}")


if __name__ == "__main__":
    main()
