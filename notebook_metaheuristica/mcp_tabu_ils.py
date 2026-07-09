"""
Búsqueda Tabú con reinicio (ILS-like) para el Maximum Clique Problem (MCP).

Construida sobre la representación y el operador de vecindad Add/Drop/Swap
especificados en Tarea 1 (tarea1_metaheuristica.tex, Secciones 2-6):

  - Representación: S subconjunto de V, codificado como máscara booleana.
  - Objetivo: f(S) = |S| (maximizar).
  - Restricción de adyacencia manejada por construcción: todo S visitado
    es, por diseño, un clique válido de G.
  - Vecindad N(S) = Add(S) U Drop(S) U Swap(S), calculada vectorizadamente
    con la matriz de adyacencia (Sección 6, Algoritmo 1 de Tarea 1).

Metaheurística: Búsqueda Tabú clásica (Glover) sobre esa vecindad, con
lista tabú de vértices recientemente removidos (no pueden re-agregarse
durante `tenure` iteraciones, salvo aspiración), y reinicio por
perturbación (estilo Búsqueda Local Iterada) cuando la búsqueda se
estanca durante `stall_limit` iteraciones sin mejorar el mejor incumbente.
"""
from __future__ import annotations

import random
import time
from dataclasses import dataclass, field

import numpy as np


def load_dimacs(path: str):
    """Lee un grafo en formato DIMACS .clq y retorna (n, A) con A matriz
    de adyacencia booleana (n x n)."""
    n = 0
    edges = []
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("c"):
                continue
            if line.startswith("p"):
                _, _, n_str, _m_str = line.split()
                n = int(n_str)
            elif line.startswith("e"):
                _, u, v = line.split()
                edges.append((int(u) - 1, int(v) - 1))  # 0-indexado
    A = np.zeros((n, n), dtype=bool)
    for u, v in edges:
        A[u, v] = True
        A[v, u] = True
    return n, A


@dataclass
class TabuILSResult:
    best_S: np.ndarray
    best_size: int
    history: list = field(default_factory=list)   # (iteracion, tiempo, f(S) actual, mejor f)
    n_restarts: int = 0
    n_iterations: int = 0
    elapsed: float = 0.0


def _candidates(A: np.ndarray, S: np.ndarray):
    """Retorna (add_candidates, swap_pairs) para la clique S.

    add_candidates: array de vértices v fuera de S adyacentes a todo S.
    swap_pairs: lista de (v, u) con v fuera de S, u en S, tal que
                intercambiar u por v mantiene S como clique.
    """
    k = int(S.sum())
    if k == 0:
        outside = np.where(~S)[0]
        return outside, []
    deg_S = A[:, S].sum(axis=1)          # deg_S[v] = |N(v) ∩ S|, para todo v
    outside_mask = ~S
    add_candidates = np.where(outside_mask & (deg_S == k))[0]
    swap_v = np.where(outside_mask & (deg_S == k - 1))[0]
    swap_pairs = []
    S_idx = np.where(S)[0]
    for v in swap_v:
        # único u en S no adyacente a v
        non_adj = S_idx[~A[v, S_idx]]
        if len(non_adj) == 1:
            swap_pairs.append((int(v), int(non_adj[0])))
    return add_candidates, swap_pairs


def _greedy_construct(A: np.ndarray, n: int, rng: random.Random, seed_S: np.ndarray | None = None):
    """Construye una clique maximal agregando vértices al azar desde C(S)
    hasta que C(S) quede vacío."""
    S = np.zeros(n, dtype=bool) if seed_S is None else seed_S.copy()
    while True:
        add_candidates, _ = _candidates(A, S)
        if len(add_candidates) == 0:
            break
        v = rng.choice(list(add_candidates))
        S[v] = True
    return S


def tabu_ils(
    A: np.ndarray,
    n: int,
    max_iterations: int = 20000,
    tenure: int = 12,
    stall_limit: int = 300,
    perturb_frac: tuple = (0.1, 0.3),
    time_limit: float | None = None,
    seed: int | None = None,
) -> TabuILSResult:
    """Búsqueda Tabú con reinicio por perturbación (ILS-like) para el MCP.

    Parámetros
    ----------
    tenure: número de iteraciones que un vértice recién eliminado
        permanece tabú para re-agregarse.
    stall_limit: iteraciones sin mejorar el mejor incumbente antes de
        aplicar una perturbación (diversificación).
    perturb_frac: fracción (min, max) de vértices de la clique actual
        que se eliminan al perturbar.
    """
    rng = random.Random(seed)
    t0 = time.time()

    S = _greedy_construct(A, n, rng)
    best_S = S.copy()
    best_size = int(S.sum())

    tabu_until = np.zeros(n, dtype=int)  # iteración hasta la cual v es tabú (para Add)
    stall = 0
    n_restarts = 0
    history = []

    it = 0
    while it < max_iterations:
        if time_limit is not None and (time.time() - t0) > time_limit:
            break
        it += 1

        add_candidates, swap_pairs = _candidates(A, S)

        if len(add_candidates) > 0:
            # Movimiento siempre mejorante: agregar (respeta tabú salvo aspiración)
            non_tabu = [v for v in add_candidates if tabu_until[v] <= it]
            pool = non_tabu if non_tabu else list(add_candidates)  # aspiración: mejora incumbente
            v = rng.choice(pool)
            S[v] = True
            stall = 0
        elif swap_pairs:
            # Movimiento lateral: swap que maximiza |C(S')| resultante (más prometedor)
            non_tabu_pairs = [(v, u) for v, u in swap_pairs if tabu_until[v] <= it]
            pool = non_tabu_pairs if non_tabu_pairs else swap_pairs
            best_pair, best_gain = None, -1
            sample = pool if len(pool) <= 15 else rng.sample(pool, 15)
            for v, u in sample:
                S_try = S.copy()
                S_try[u] = False
                S_try[v] = True
                add_after, _ = _candidates(A, S_try)
                gain = len(add_after)
                if gain > best_gain:
                    best_gain, best_pair = gain, (v, u)
            v, u = best_pair
            S[u] = False
            S[v] = True
            tabu_until[u] = it + tenure
            stall += 1
        else:
            # Clique maximal sin swaps disponibles: drop forzado (diversifica)
            S_idx = np.where(S)[0]
            if len(S_idx) == 0:
                S = _greedy_construct(A, n, rng)
                continue
            u = rng.choice(list(S_idx))
            S[u] = False
            tabu_until[u] = it + tenure
            stall += 1

        cur_size = int(S.sum())
        if cur_size > best_size:
            best_size = cur_size
            best_S = S.copy()
            stall = 0

        history.append((it, time.time() - t0, cur_size, best_size))

        if stall >= stall_limit:
            # Perturbación (reinicio ILS): eliminar una fracción aleatoria
            # de la clique actual y reconstruir de forma golosa.
            n_restarts += 1
            S_idx = np.where(S)[0]
            frac = rng.uniform(*perturb_frac)
            n_remove = max(1, int(len(S_idx) * frac))
            to_remove = rng.sample(list(S_idx), min(n_remove, len(S_idx)))
            for u in to_remove:
                S[u] = False
                tabu_until[u] = it + tenure
            S = _greedy_construct(A, n, rng, seed_S=S)
            stall = 0

    elapsed = time.time() - t0
    return TabuILSResult(
        best_S=best_S,
        best_size=best_size,
        history=history,
        n_restarts=n_restarts,
        n_iterations=it,
        elapsed=elapsed,
    )


def verify_clique(A: np.ndarray, S: np.ndarray) -> bool:
    idx = np.where(S)[0]
    k = len(idx)
    if k <= 1:
        return True
    sub = A[np.ix_(idx, idx)]
    n_edges = sub.sum() // 2
    return n_edges == k * (k - 1) // 2
