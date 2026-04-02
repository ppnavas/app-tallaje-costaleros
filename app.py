import streamlit as st
import pandas as pd
import io
from datetime import datetime
from itertools import combinations as icombs, permutations as iperms
import random, math, time

def hard_ok(c, lado):
    """Devuelve False si la preferencia estricta prohíbe este lado."""
    p = c['Preferencia de Hombro']
    return (
        not (p == 'Solo Izquierdo' and lado == 'Izquierdo') and
        not (p == 'Solo Derecho'   and lado == 'Derecho')
    )

def run_assignment_algorithm(pool, varales_config, capacidad_ext, capacidad_int, num_varales, max_filas,
                             custom_weights=None, relax_cross_row=False, seed_offset=0,
                             progress_callback=None):
    """Ejecuta el pipeline completo: Fase 1 greedy, 1b reparación, 2 SA, 3 hill climbing."""

    asignaciones_por_varal = {v["Nombre"]: [] for v in varales_config}

    # ── FUNCIONES AUXILIARES ───────────────────────────────────────────

    def h_carga(c, lado):
        """Altura del hombro que realmente carga según el lado del varal."""
        return (
            c['Altura Hombro Derecho (cm)']
            if lado == 'Izquierdo'
            else c['Altura Hombro Izquierdo (cm)']
        )

    def pref_sat(c, lado):
        """True si la preferencia blanda del costalero se satisface en este lado."""
        p = c['Preferencia de Hombro']
        return (
            (p == 'Derecho'        and lado == 'Izquierdo') or
            (p == 'Izquierdo'      and lado == 'Derecho')   or
            (p == 'Solo Derecho'   and lado == 'Izquierdo') or
            (p == 'Solo Izquierdo' and lado == 'Derecho')   or
            p == 'Indiferente'
        )

    def fila_activa(fila, tipo):
        """True si este tipo de varal participa en esta fila."""
        if tipo == "Exterior":
            return fila <= capacidad_ext
        return (
            fila <= capacidad_int // 2 or
            fila >  max_filas - capacidad_int // 2
        )

    # ── ESTRUCTURA DE POSICIONES Y FUNCIÓN OBJETIVO ────────────────────

    # Construir lista de todas las posiciones y sus emparejamientos
    all_positions = []   # [(varal_name, row_idx, lado, tipo)]
    varal_filas = {}     # varal_name -> [fila1, fila2, ...]
    for v in varales_config:
        filas_v = [f for f in range(1, max_filas + 1) if fila_activa(f, v["Tipo"])]
        varal_filas[v['Nombre']] = filas_v
        for idx in range(len(filas_v)):
            all_positions.append((v['Nombre'], idx, v['Lado'], v['Tipo']))

    # Emparejar posiciones izq/der del mismo tipo en la misma fila
    pair_list = []  # [(varal_izq, varal_der, row_idx, fila)]
    left_varals  = [v for v in varales_config if v['Lado'] == 'Izquierdo']
    right_varals = [v for v in varales_config if v['Lado'] == 'Derecho']
    for vl in left_varals:
        for vr in right_varals:
            if vl['Tipo'] == vr['Tipo']:
                filas_l = varal_filas[vl['Nombre']]
                filas_r = varal_filas[vr['Nombre']]
                for idx in range(min(len(filas_l), len(filas_r))):
                    pair_list.append((vl['Nombre'], vr['Nombre'], idx, filas_l[idx]))

    # Mapa fila -> lista de (varal_name, row_idx, lado)
    fila_positions = {}
    for vname, idx, lado, tipo in all_positions:
        f = varal_filas[vname][idx]
        fila_positions.setdefault(f, []).append((vname, idx, lado))

    # Lado y tipo de cada varal (cache)
    varal_lado = {v['Nombre']: v['Lado'] for v in varales_config}
    varal_tipo = {v['Nombre']: v['Tipo'] for v in varales_config}

    # Filas consecutivas para verificar orden entre filas
    filas_ordenadas = sorted(fila_positions.keys())

    W_PAIR  = 15.0
    W_ROW   = 10.0
    W_PREF  = 2.0
    W_GRAD  = 5.0
    W_CROSS = 50.0  # Backup de la restricción dura cross-row
    W_EXT   = 5.0   # Exteriores = más altos de la fila
    if custom_weights:
        W_PAIR  = custom_weights.get('W_PAIR', W_PAIR)
        W_ROW   = custom_weights.get('W_ROW', W_ROW)
        W_PREF  = custom_weights.get('W_PREF', W_PREF)
        W_GRAD  = custom_weights.get('W_GRAD', W_GRAD)
        W_CROSS = custom_weights.get('W_CROSS', W_CROSS)
        W_EXT   = custom_weights.get('W_EXT', W_EXT)

    def compute_J(grid):
        """Función objetivo global. Menor = mejor."""
        # J_pair: diferencia^2 en cada par de posiciones enfrentadas
        j_pair = 0.0
        for vl, vr, idx, fila in pair_list:
            cl, cr = grid[vl][idx], grid[vr][idx]
            diff = h_carga(cl, 'Izquierdo') - h_carga(cr, 'Derecho')
            j_pair += diff * diff

        # J_row: varianza de alturas dentro de cada fila
        j_row = 0.0
        for f, pos_list in fila_positions.items():
            heights = [h_carga(grid[vn][ri], lado) for vn, ri, lado in pos_list]
            if len(heights) > 1:
                mean_h = sum(heights) / len(heights)
                j_row += sum((h - mean_h) ** 2 for h in heights) / len(heights)

        # J_pref: preferencias blandas no satisfechas
        j_pref = 0
        for vname, idx, lado, tipo in all_positions:
            if not pref_sat(grid[vname][idx], lado):
                j_pref += 1

        # J_grad: suavidad de alturas dentro de cada varal
        j_grad = 0.0
        for v in varales_config:
            col = grid[v['Nombre']]
            lado = v['Lado']
            filas_v = varal_filas[v['Nombre']]
            for i in range(len(col) - 1):
                if filas_v[i + 1] - filas_v[i] != 1:
                    continue
                h_curr = h_carga(col[i], lado)
                h_next = h_carga(col[i + 1], lado)
                diff = h_next - h_curr
                if diff > 0:
                    j_grad += diff ** 2 * 10
                else:
                    j_grad += diff ** 2

        # J_cross: penalización por inversiones ENTRE filas (cross-varal)
        j_cross = 0.0
        for ri in range(len(filas_ordenadas) - 1):
            f_curr = filas_ordenadas[ri]
            f_next = filas_ordenadas[ri + 1]
            min_curr = min(h_carga(grid[vn][ri_idx], lado)
                        for vn, ri_idx, lado in fila_positions[f_curr])
            max_next = max(h_carga(grid[vn][ri_idx], lado)
                        for vn, ri_idx, lado in fila_positions[f_next])
            if max_next > min_curr:
                j_cross += (max_next - min_curr) ** 2

        # J_ext: exteriores deben tener los costaleros más altos de cada fila
        j_ext = 0.0
        for f, pos_list in fila_positions.items():
            for side in ('Izquierdo', 'Derecho'):
                ext_h = [h_carga(grid[vn][ri], lado)
                        for vn, ri, lado in pos_list
                        if lado == side and varal_tipo[vn] == 'Exterior']
                int_h = [h_carga(grid[vn][ri], lado)
                        for vn, ri, lado in pos_list
                        if lado == side and varal_tipo[vn] == 'Interior']
                if ext_h and int_h:
                    viol = max(int_h) - min(ext_h)
                    if viol > 0:
                        j_ext += viol ** 2

        return (W_PAIR * j_pair + W_ROW * j_row + W_PREF * j_pref
                + W_GRAD * j_grad + W_CROSS * j_cross + W_EXT * j_ext)

    def check_swap_ok(grid, v1, i1, lado1, v2, i2, lado2):
        """True si el intercambio es factible (restricciones duras)."""
        c1, c2 = grid[v1][i1], grid[v2][i2]

        if not hard_ok(c2, lado1) or not hard_ok(c1, lado2):
            return False

        if relax_cross_row:
            return True

        h1_new = h_carga(c2, lado1)
        h2_new = h_carga(c1, lado2)

        fila1 = varal_filas[v1][i1]
        fila2 = varal_filas[v2][i2]
        affected = {fila1, fila2}
        for ri in range(len(filas_ordenadas) - 1):
            f_c = filas_ordenadas[ri]
            f_n = filas_ordenadas[ri + 1]
            if f_c not in affected and f_n not in affected:
                continue
            heights_c, heights_n = [], []
            for vn, ri_idx, lado in fila_positions[f_c]:
                if vn == v1 and ri_idx == i1:
                    heights_c.append(h1_new)
                elif vn == v2 and ri_idx == i2:
                    heights_c.append(h2_new)
                else:
                    heights_c.append(h_carga(grid[vn][ri_idx], lado))
            for vn, ri_idx, lado in fila_positions[f_n]:
                if vn == v1 and ri_idx == i1:
                    heights_n.append(h1_new)
                elif vn == v2 and ri_idx == i2:
                    heights_n.append(h2_new)
                else:
                    heights_n.append(h_carga(grid[vn][ri_idx], lado))
            if max(heights_n) > min(heights_c):
                return False

        return True

    def check_rotate_ok(grid, positions_3):
        """True si la rotación A->B->C->A es factible."""
        (v1, i1, l1), (v2, i2, l2), (v3, i3, l3) = positions_3
        c1, c2, c3 = grid[v1][i1], grid[v2][i2], grid[v3][i3]
        if not hard_ok(c3, l1) or not hard_ok(c1, l2) or not hard_ok(c2, l3):
            return False

        if relax_cross_row:
            return True

        rotate_map = {(v1, i1): c3, (v2, i2): c1, (v3, i3): c2}
        affected_r = {varal_filas[v1][i1], varal_filas[v2][i2], varal_filas[v3][i3]}
        for ri in range(len(filas_ordenadas) - 1):
            f_c = filas_ordenadas[ri]
            f_n = filas_ordenadas[ri + 1]
            if f_c not in affected_r and f_n not in affected_r:
                continue

            def sim_h(vn, ri_idx, lado):
                key = (vn, ri_idx)
                if key in rotate_map:
                    return h_carga(rotate_map[key], lado)
                return h_carga(grid[vn][ri_idx], lado)

            min_c = min(sim_h(vn, ri_idx, lado)
                        for vn, ri_idx, lado in fila_positions[f_c])
            max_n = max(sim_h(vn, ri_idx, lado)
                        for vn, ri_idx, lado in fila_positions[f_n])
            if max_n > min_c:
                return False

        return True

    # ── ASIGNACIÓN ÓPTIMA POR FILA (FASE 1) ─────────────────────────────

    def asignar_fila(candidatos, huecos_fila):
        """
        Asigna los candidatos a los huecos de la fila de forma óptima.
        """
        huecos_izq = sorted(
            [h for h in huecos_fila if h['Lado'] == 'Izquierdo'],
            key=lambda h: 0 if h['Tipo'] == 'Exterior' else 1
        )
        huecos_der = sorted(
            [h for h in huecos_fila if h['Lado'] == 'Derecho'],
            key=lambda h: 0 if h['Tipo'] == 'Exterior' else 1
        )
        n_izq = len(huecos_izq)

        forzados_izq = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Derecho']
        forzados_der = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Izquierdo']
        libres       = [c for c in candidatos
                        if c['Preferencia de Hombro'] not in ('Solo Izquierdo', 'Solo Derecho')]
        n_flex_izq   = n_izq - len(forzados_izq)

        if not (0 <= n_flex_izq <= len(libres)):
            res_fallback = {}
            used = set()
            for h in huecos_fila:
                best_c, best_idx = None, -1
                for ci, c in enumerate(candidatos):
                    if ci in used:
                        continue
                    if hard_ok(c, h['Lado']):
                        best_c, best_idx = c, ci
                        break
                if best_c is None:
                    for ci, c in enumerate(candidatos):
                        if ci not in used:
                            best_c, best_idx = c, ci
                            break
                res_fallback[h['Nombre']] = best_c
                used.add(best_idx)
            return res_fallback

        best_score = (float('inf'), float('inf'), float('inf'))
        best_res   = None

        for sel in icombs(range(len(libres)), n_flex_izq):
            sel_set = set(sel)
            g_izq = forzados_izq + [libres[i] for i in sel]
            g_der = forzados_der + [libres[i] for i in range(len(libres)) if i not in sel_set]

            for perm_i in iperms(g_izq):
                for perm_d in iperms(g_der):
                    pair_score = sum(
                        (h_carga(gl, 'Izquierdo') - h_carga(gr, 'Derecho')) ** 2
                        for gl, gr in zip(perm_i, perm_d)
                    )

                    ext_penalty = 0.0
                    for huecos_s, perm_s, lado_s in [
                        (huecos_izq, perm_i, 'Izquierdo'),
                        (huecos_der, perm_d, 'Derecho')
                    ]:
                        e_h = [h_carga(c, lado_s) for h, c in zip(huecos_s, perm_s)
                            if h['Tipo'] == 'Exterior']
                        i_h = [h_carga(c, lado_s) for h, c in zip(huecos_s, perm_s)
                            if h['Tipo'] == 'Interior']
                        if e_h and i_h:
                            v = max(i_h) - min(e_h)
                            if v > 0:
                                ext_penalty += v ** 2

                    prefs = (
                        sum(pref_sat(c, 'Izquierdo') for c in perm_i) +
                        sum(pref_sat(c, 'Derecho')   for c in perm_d)
                    )

                    score = (pair_score, ext_penalty, -prefs)
                    if score < best_score:
                        best_score = score
                        best_res   = {}
                        for h, c in zip(huecos_izq, perm_i): best_res[h['Nombre']] = c
                        for h, c in zip(huecos_der, perm_d): best_res[h['Nombre']] = c

        return best_res or {h['Nombre']: candidatos[i] for i, h in enumerate(huecos_fila)}

    # ── FASE 1 — ASIGNACIÓN FILA A FILA ──────────────────────────────────

    pool = list(pool)  # copia local
    for fila in range(1, max_filas + 1):
        huecos_fila = [v for v in varales_config if fila_activa(fila, v["Tipo"])]
        if not huecos_fila:
            continue

        req_total = len(huecos_fila)
        req_izq   = sum(1 for h in huecos_fila if h['Lado'] == 'Izquierdo')
        req_der   = req_total - req_izq

        candidatos, pool_restante, temp_pool = [], [], pool.copy()
        while len(candidatos) < req_total and temp_pool:
            necesarios = req_total - len(candidatos)
            candidatos.extend(temp_pool[:necesarios])
            temp_pool = temp_pool[necesarios:]

            solo_der = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Derecho']
            solo_izq = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Izquierdo']

            while len(solo_der) > req_izq:
                peor = solo_der.pop(-1)
                candidatos.remove(peor)
                pool_restante.append(peor)

            while len(solo_izq) > req_der:
                peor = solo_izq.pop(-1)
                candidatos.remove(peor)
                pool_restante.append(peor)

        # Fallback: si temp_pool se agotó sin suficientes candidatos, tomar de pool_restante
        if len(candidatos) < req_total:
            pool_restante.sort(key=lambda x: x['Altura_Media'], reverse=True)
            faltantes = req_total - len(candidatos)
            candidatos.extend(pool_restante[:faltantes])
            pool_restante = pool_restante[faltantes:]

        pool = sorted(
            pool_restante + temp_pool,
            key=lambda x: x['Altura_Media'], reverse=True
        )

        res = asignar_fila(candidatos, huecos_fila)
        for v in huecos_fila:
            asignaciones_por_varal[v['Nombre']].append(res[v['Nombre']])

    # ── REPARACIÓN CROSS-ROW ─────────────────────────────────────────

    grid = {vn: list(cs) for vn, cs in asignaciones_por_varal.items()}

    for _repair_iter in range(100):
        violation_found = False
        for ri in range(len(filas_ordenadas) - 1):
            f_c = filas_ordenadas[ri]
            f_n = filas_ordenadas[ri + 1]
            pos_c = fila_positions[f_c]
            pos_n = fila_positions[f_n]
            h_c = [h_carga(grid[vn][idx], lado) for vn, idx, lado in pos_c]
            h_n = [h_carga(grid[vn][idx], lado) for vn, idx, lado in pos_n]

            if max(h_n) <= min(h_c):
                continue

            violation_found = True
            best_swap = None
            best_violation = max(h_n) - min(h_c)
            for a, (vn_a, idx_a, lado_a) in enumerate(pos_c):
                for b, (vn_b, idx_b, lado_b) in enumerate(pos_n):
                    ca = grid[vn_a][idx_a]
                    cb = grid[vn_b][idx_b]
                    if not hard_ok(cb, lado_a) or not hard_ok(ca, lado_b):
                        continue
                    new_h_c = list(h_c)
                    new_h_c[a] = h_carga(cb, lado_a)
                    new_h_n = list(h_n)
                    new_h_n[b] = h_carga(ca, lado_b)
                    new_viol = max(max(new_h_n) - min(new_h_c), 0)
                    if new_viol < best_violation:
                        best_violation = new_viol
                        best_swap = (vn_a, idx_a, vn_b, idx_b)

            if best_swap:
                va, ia, vb, ib = best_swap
                grid[va][ia], grid[vb][ib] = grid[vb][ib], grid[va][ia]

        if not violation_found:
            break

    # ── FASE 2 — SIMULATED ANNEALING ─────────────────────────────────

    N_pos = len(all_positions)

    same_row_same_type = []
    same_row_diff_type = []
    same_varal_pairs = []
    cross_pairs = []
    same_side_same_row = []

    by_fila = {}
    for pi, (vn, idx, lado, tipo) in enumerate(all_positions):
        f = varal_filas[vn][idx]
        by_fila.setdefault(f, []).append(pi)

    for f, pis in by_fila.items():
        for a in range(len(pis)):
            for b in range(a + 1, len(pis)):
                pa, pb = all_positions[pis[a]], all_positions[pis[b]]
                if pa[2] != pb[2]:
                    if pa[3] == pb[3]:
                        same_row_same_type.append((pis[a], pis[b]))
                    else:
                        same_row_diff_type.append((pis[a], pis[b]))
                elif pa[0] != pb[0]:
                    same_side_same_row.append((pis[a], pis[b]))

    vname_to_indices = {}
    for pi, (vn, idx, lado, tipo) in enumerate(all_positions):
        vname_to_indices.setdefault(vn, []).append(pi)
    for vn, pis in vname_to_indices.items():
        for a in range(len(pis)):
            for b in range(a + 1, len(pis)):
                same_varal_pairs.append((pis[a], pis[b]))

    filas_sorted = sorted(by_fila.keys())
    for fi_idx, f1 in enumerate(filas_sorted):
        for f2 in filas_sorted[fi_idx + 1:]:
            for a in by_fila[f1]:
                for b in by_fila[f2]:
                    if all_positions[a][2] != all_positions[b][2]:
                        cross_pairs.append((a, b))

    def do_swap(grid, pi_a, pi_b):
        va, ia, la, _ = all_positions[pi_a]
        vb, ib, lb, _ = all_positions[pi_b]
        grid[va][ia], grid[vb][ib] = grid[vb][ib], grid[va][ia]

    def copy_grid(g):
        return {vn: list(cs) for vn, cs in g.items()}

    best_grid = copy_grid(grid)
    best_J = compute_J(grid)

    SA_RESTARTS = 3
    T0 = 10.0
    T_MIN = 0.001
    ALPHA = 0.97
    ITERS_PER_T = max(N_pos * 5, 50)

    for restart in range(SA_RESTARTS):
        if restart > 0:
            grid = copy_grid(best_grid)

        current_J = compute_J(grid)
        rng = random.Random(42 + restart * 1000 + seed_offset * 100000)
        T = T0

        while T > T_MIN:
            for _ in range(ITERS_PER_T):
                r = rng.random()

                pi_a = pi_b = -1
                is_rotate = False

                if r < 0.35 and same_row_same_type:
                    pi_a, pi_b = rng.choice(same_row_same_type)
                elif r < 0.45 and same_row_diff_type:
                    pi_a, pi_b = rng.choice(same_row_diff_type)
                elif r < 0.65 and same_varal_pairs:
                    pi_a, pi_b = rng.choice(same_varal_pairs)
                elif r < 0.75 and cross_pairs:
                    pi_a, pi_b = rng.choice(cross_pairs)
                elif r < 0.95 and same_side_same_row:
                    pi_a, pi_b = rng.choice(same_side_same_row)
                elif N_pos >= 3:
                    is_rotate = True
                    tri = rng.sample(range(N_pos), 3)
                    p0 = (all_positions[tri[0]][0], all_positions[tri[0]][1], all_positions[tri[0]][2])
                    p1 = (all_positions[tri[1]][0], all_positions[tri[1]][1], all_positions[tri[1]][2])
                    p2 = (all_positions[tri[2]][0], all_positions[tri[2]][1], all_positions[tri[2]][2])
                else:
                    continue

                if is_rotate:
                    if not check_rotate_ok(grid, (p0, p1, p2)):
                        continue
                    (v0, i0, _), (v1, i1, _), (v2, i2, _) = p0, p1, p2
                    old_c0, old_c1, old_c2 = grid[v0][i0], grid[v1][i1], grid[v2][i2]
                    grid[v0][i0], grid[v1][i1], grid[v2][i2] = old_c2, old_c0, old_c1
                    new_J = compute_J(grid)
                    delta = new_J - current_J
                    if delta < 0 or rng.random() < math.exp(-delta / T):
                        current_J = new_J
                    else:
                        grid[v0][i0], grid[v1][i1], grid[v2][i2] = old_c0, old_c1, old_c2
                else:
                    if pi_a < 0:
                        continue
                    va, ia, la, _ = all_positions[pi_a]
                    vb, ib, lb, _ = all_positions[pi_b]
                    if not check_swap_ok(grid, va, ia, la, vb, ib, lb):
                        continue
                    do_swap(grid, pi_a, pi_b)
                    new_J = compute_J(grid)
                    delta = new_J - current_J
                    if delta < 0 or rng.random() < math.exp(-delta / T):
                        current_J = new_J
                    else:
                        do_swap(grid, pi_a, pi_b)

            T *= ALPHA

        if current_J < best_J:
            best_J = current_J
            best_grid = copy_grid(grid)

        if progress_callback:
            progress_callback(restart + 1, SA_RESTARTS)

    grid = best_grid

    # ── FASE 3 — HILL CLIMBING DETERMINISTA ──────────────────────────

    current_J = compute_J(grid)
    improved = True
    while improved:
        improved = False
        for i in range(N_pos):
            for j in range(i + 1, N_pos):
                vi, ii, li, _ = all_positions[i]
                vj, ij, lj, _ = all_positions[j]
                if not check_swap_ok(grid, vi, ii, li, vj, ij, lj):
                    continue
                do_swap(grid, i, j)
                new_J = compute_J(grid)
                if new_J < current_J - 0.001:
                    current_J = new_J
                    improved = True
                else:
                    do_swap(grid, i, j)

    # ── GARANTÍA CROSS-ROW (safety net) ─────────────────────────────
    def _main_cross_cost(g):
        total = 0.0
        for ri in range(len(filas_ordenadas) - 1):
            fc, fn = filas_ordenadas[ri], filas_ordenadas[ri + 1]
            mn = min(h_carga(g[vn][idx], lado) for vn, idx, lado in fila_positions[fc])
            mx = max(h_carga(g[vn][idx], lado) for vn, idx, lado in fila_positions[fn])
            if mx > mn:
                total += (mx - mn)
        return total

    # Fase A: rotaciones de 3 vías
    if _main_cross_cost(grid) > 0.001:
        all_main_pos = []
        for f in filas_ordenadas:
            for vn, idx, lado in fila_positions[f]:
                all_main_pos.append((vn, idx, lado, f))

        for _rot in range(200):
            cost_before = _main_cross_cost(grid)
            if cost_before < 0.001:
                break
            improved = False
            for i in range(len(all_main_pos)):
                if improved:
                    break
                for j in range(len(all_main_pos)):
                    if improved:
                        break
                    if i == j:
                        continue
                    for k in range(len(all_main_pos)):
                        if k == i or k == j:
                            continue
                        vn_i, idx_i, _, _ = all_main_pos[i]
                        vn_j, idx_j, _, _ = all_main_pos[j]
                        vn_k, idx_k, _, _ = all_main_pos[k]
                        ci_v = grid[vn_i][idx_i]
                        cj_v = grid[vn_j][idx_j]
                        ck_v = grid[vn_k][idx_k]
                        _, _, l_i, _ = all_main_pos[i]
                        _, _, l_j, _ = all_main_pos[j]
                        _, _, l_k, _ = all_main_pos[k]
                        if not hard_ok(cj_v, l_i) or not hard_ok(ck_v, l_j) or not hard_ok(ci_v, l_k):
                            continue
                        grid[vn_i][idx_i] = cj_v
                        grid[vn_j][idx_j] = ck_v
                        grid[vn_k][idx_k] = ci_v
                        if _main_cross_cost(grid) < cost_before - 0.001:
                            improved = True
                            break
                        grid[vn_i][idx_i] = ci_v
                        grid[vn_j][idx_j] = cj_v
                        grid[vn_k][idx_k] = ck_v
            if not improved:
                break

    # Fase B: redistribución global por Altura_Media + bubble sort
    if _main_cross_cost(grid) > 0.001:
        grid_backup = {vn: list(cs) for vn, cs in grid.items()}

        all_c = []
        for f in filas_ordenadas:
            for vn, idx, lado in fila_positions[f]:
                all_c.append(grid[vn][idx])
        all_c.sort(key=lambda c: c['Altura_Media'], reverse=True)

        ci_pos = 0
        hard_violated = False
        for f in filas_ordenadas:
            positions = fila_positions[f]
            n = len(positions)
            row_c = all_c[ci_pos:ci_pos + n]
            ci_pos += n
            used = [False] * n
            for vn, idx, lado in positions:
                best_j, best_h = -1, -1.0
                for j in range(n):
                    if used[j]:
                        continue
                    if not hard_ok(row_c[j], lado):
                        continue
                    h = h_carga(row_c[j], lado)
                    if h > best_h:
                        best_h, best_j = h, j
                if best_j == -1:
                    hard_violated = True
                    for j in range(n):
                        if used[j]:
                            continue
                        h = h_carga(row_c[j], lado)
                        if h > best_h:
                            best_h, best_j = h, j
                grid[vn][idx] = row_c[best_j]
                used[best_j] = True

        if hard_violated:
            grid = grid_backup
        else:
            for _bs in range(1000):
                if _main_cross_cost(grid) < 0.001:
                    break
                swapped = False
                for ri in range(len(filas_ordenadas) - 1):
                    fc = filas_ordenadas[ri]
                    fn = filas_ordenadas[ri + 1]
                    mn = min(h_carga(grid[vn][idx], lado) for vn, idx, lado in fila_positions[fc])
                    mx = max(h_carga(grid[vn][idx], lado) for vn, idx, lado in fila_positions[fn])
                    if mx <= mn:
                        continue
                    pos_short = min(fila_positions[fc], key=lambda t: h_carga(grid[t[0]][t[1]], t[2]))
                    pos_tall = max(fila_positions[fn], key=lambda t: h_carga(grid[t[0]][t[1]], t[2]))
                    vn_a, idx_a, lado_a = pos_short
                    vn_b, idx_b, lado_b = pos_tall
                    if hard_ok(grid[vn_b][idx_b], lado_a) and hard_ok(grid[vn_a][idx_a], lado_b):
                        grid[vn_a][idx_a], grid[vn_b][idx_b] = grid[vn_b][idx_b], grid[vn_a][idx_a]
                        swapped = True
                if not swapped:
                    break

    # Fase C: greedy global — probar TODOS los pares del grid
    if _main_cross_cost(grid) > 0.001:
        all_main_pos = []
        for f in filas_ordenadas:
            for vn, idx, lado in fila_positions[f]:
                all_main_pos.append((vn, idx, lado, f))

        for _ in range(500):
            cc = _main_cross_cost(grid)
            if cc < 0.001:
                break
            best_swap, best_cost = None, cc
            for i in range(len(all_main_pos)):
                for j in range(i + 1, len(all_main_pos)):
                    vn_a, idx_a, la, fa = all_main_pos[i]
                    vn_b, idx_b, lb, fb = all_main_pos[j]
                    if fa == fb:
                        continue
                    ca, cb = grid[vn_a][idx_a], grid[vn_b][idx_b]
                    if not hard_ok(cb, la) or not hard_ok(ca, lb):
                        continue
                    grid[vn_a][idx_a], grid[vn_b][idx_b] = cb, ca
                    nc = _main_cross_cost(grid)
                    if nc < best_cost - 0.001:
                        best_cost = nc
                        best_swap = (vn_a, idx_a, vn_b, idx_b)
                    grid[vn_a][idx_a], grid[vn_b][idx_b] = ca, cb
            if not best_swap:
                break
            va, ia, vb, ib = best_swap
            grid[va][ia], grid[vb][ib] = grid[vb][ib], grid[va][ia]


    # Copiar resultado al diccionario original
    for vn in asignaciones_por_varal:
        asignaciones_por_varal[vn] = grid[vn]

    return (asignaciones_por_varal, grid, all_positions, varal_filas, fila_positions,
            fila_activa, h_carga, varal_lado)

def reset_analisis():
    st.session_state.analisis_completado = False
    st.session_state.excel_buffer = None
    st.session_state.num_asignados = 0
    st.session_state.total_titulares = 0
    st.session_state.total_suplentes = 0
    st.session_state.df_titulares = None
    st.session_state.df_suplentes = None
    st.session_state.last_max = 0

# 1. Detección del dispositivo
ua = st.context.headers.get("User-Agent", "").lower()
es_movil = any(x in ua for x in ["iphone", "android", "mobile", "ipad"])

# Configuración de página
st.set_page_config(page_title="Tallaje de Costaleros", page_icon="logo.png", layout="wide", initial_sidebar_state="collapsed")

año_actual = datetime.now().year

st.title("Tallaje de Costaleros")
st.divider()

# --- 1. ESTADO DE LA SESIÓN ---
if "analisis_completado" not in st.session_state: st.session_state.analisis_completado = False
if "excel_buffer" not in st.session_state: st.session_state.excel_buffer = None
if "num_asignados" not in st.session_state: st.session_state.num_asignados = 0
if "total_titulares" not in st.session_state: st.session_state.total_titulares = 0
if "total_suplentes" not in st.session_state: st.session_state.total_suplentes = 0
if "df_titulares" not in st.session_state: st.session_state.df_titulares = None
if "df_suplentes" not in st.session_state: st.session_state.df_suplentes = None

if es_movil:
    col1 = st.container()
    st.divider()
    col2 = st.container()
    st.divider()
    col3 = st.container()
else:
    col1, col2, col3 = st.columns(3, gap="large")

# ==========================================
# COLUMNA 1: MAX COSTALEROS Y PLANTILLA
# ==========================================
with col1:

    st.markdown("#### Plantilla")
    st.write("")

    st.markdown("###### Número máximo de costaleros")
    val_act_max = st.session_state.get("max_port", 60)
    paso_max = 1 if val_act_max % 2 != 0 else 2

    max_portadores = st.number_input(
        "Número máximo de costaleros", 
        min_value=4, 
        value=60, 
        step=paso_max, 
        key="max_port", 
        label_visibility="collapsed"
    )
    
    st.write("")    
    
    # Condición para cambiar el mensaje
    if max_portadores % 2 != 0:
        st.error("✖ &nbsp;&nbsp;&nbsp;El número máximo de costaleros debe ser par.")
    else:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Descarga esta plantilla y rellénala con los datos de cada costalero.")

    df_plantilla = pd.DataFrame(columns=[
        "Nombre", "Preferencia de Hombro", "Altura Hombro Izquierdo (cm)", "Altura Hombro Derecho (cm)"
    ])
    
    buffer_plantilla = io.BytesIO()
    with pd.ExcelWriter(buffer_plantilla, engine='xlsxwriter') as writer:
        pestana_1 = 'Titulares' 
        pestana_2 = 'Suplentes'
        
        df_plantilla.to_excel(writer, index=False, sheet_name=pestana_1, startrow=1)
        df_plantilla.to_excel(writer, index=False, sheet_name=pestana_2, startrow=1)
        
        titulo_fmt = writer.book.add_format({'bold': True, 'bg_color': '#FF2B2B', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        titulo_supl_fmt = writer.book.add_format({'bold': True, 'bg_color': "#FF2B2B", 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        header_fmt = writer.book.add_format({'bold': True, 'bg_color': "#8583FF", 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 13})
        cell_fmt = writer.book.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 12}) 
        
        num_filas = max_portadores + 2 
        opciones_hombro = ['Indiferente', 'Izquierdo', 'Derecho', 'Solo Izquierdo', 'Solo Derecho']

        configuracion = [
            (pestana_1, 'Costaleros Titulares', titulo_fmt), 
            (pestana_2, 'Costaleros Suplentes', titulo_supl_fmt)
        ]

        for nombre_pestana, titulo_fila_1, formato_titulo in configuracion:
            ws = writer.sheets[nombre_pestana]
            
            ws.set_row(0, 35)
            ws.merge_range(0, 0, 0, 3, titulo_fila_1, formato_titulo)

            ws.set_row(1, 30)
            for i, col in enumerate(df_plantilla.columns):
                ws.set_column(i, i, 40)
                ws.write(1, i, col, header_fmt)

            for fila in range(2, num_filas):
                ws.set_row(fila, 25) 
                for col in range(4): 
                    ws.write_blank(fila, col, "", cell_fmt)
            
            ws.data_validation(2, 1, num_filas - 1, 1, {
                'validate': 'list', 'source': opciones_hombro,
                'error_title': 'Opción no válida', 'error_message': 'Por favor, elige una opción de la lista.'
            })
            ws.data_validation(2, 2, num_filas, 3, {
                'validate': 'decimal', 'criteria': '>', 'value': 0,
                'error_title': 'Dato no válido', 'error_message': 'Introduce un valor numérico.'
            })

    st.download_button(
        label="Descargar", data=buffer_plantilla.getvalue(), 
        file_name="plantilla_costaleros.xlsx", use_container_width=True,
        disabled=(max_portadores % 2 != 0) # Aquí se bloquea si es impar
    )

    st.divider()

    st.markdown("#### Carga de datos")
    st.write("")

    st.info("ⓘ &nbsp;&nbsp;&nbsp;Sube aquí la plantilla rellenada.")
    
    archivo = st.file_uploader("Sube el Excel", type=["xlsx"], label_visibility="collapsed", on_change=reset_analisis)

    # Leemos el Excel automáticamente al subirlo para detectar cantidades
    # dropna(subset=['Nombre']) asegura que solo se cuenten las filas donde se ha escrito un nombre
    if archivo is not None and st.session_state.df_titulares is None:
        try:
            df_t = pd.read_excel(archivo, sheet_name='Titulares', skiprows=1).dropna(subset=['Nombre'])
            df_s = pd.read_excel(archivo, sheet_name='Suplentes', skiprows=1).dropna(subset=['Nombre'])
            st.session_state.total_titulares = len(df_t)
            st.session_state.total_suplentes = len(df_s)
            st.session_state.df_titulares = df_t
            st.session_state.df_suplentes = df_s
        except:
            st.session_state.total_titulares = 0

    

# ==========================================
# COLUMNA 2: ARCHIVO -> PARÁMETROS -> ANÁLISIS
# ==========================================
with col2:
    st.markdown("#### Parámetros generales")
    st.write("")

    if archivo is not None:
        st.write(f"Costaleros detectados &nbsp; ⟶ &nbsp; Titulares: **{st.session_state.total_titulares}** &nbsp;|&nbsp; Suplentes: **{st.session_state.total_suplentes}**")
        st.write("")

    total_t = st.session_state.total_titulares
    parametros_validos = False
    varales_config = []

    # BARRERA: Solo mostramos la configuración si hay archivo subido y detecta costaleros
    if st.session_state.df_titulares is None:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Carga los datos para acceder a la configuración.")
    elif total_t == 0:
        st.error("⚠ &nbsp;&nbsp;&nbsp;No se han detectado costaleros titulares en el archivo.")
    elif total_t % 2 != 0:
        st.error(f"⚠ &nbsp;&nbsp;&nbsp;El número de titulares ({total_t}) debe ser par.")
    else:
        def actualizar_desde_interior():
            st.session_state.err_int = False
            st.session_state.err_ext = False
            c_int = st.session_state.cap_int_edit
            n_var = st.session_state.num_var
            max_p = st.session_state.total_titulares
            c_ext = (max_p - (n_var - 2) * c_int) // 2
            if c_ext >= 2: st.session_state.cap_ext_edit = c_ext
            else: st.session_state.err_int = True

        def actualizar_desde_exterior():
            st.session_state.err_int = False
            st.session_state.err_ext = False
            c_ext = st.session_state.cap_ext_edit
            n_var = st.session_state.num_var
            max_p = st.session_state.total_titulares
            if n_var > 2:
                resto = max_p - 2 * c_ext
                c_int = resto / (n_var - 2)
                if c_int >= 2 and c_int % 2 == 0 and c_int.is_integer(): st.session_state.cap_int_edit = int(c_int)
                else: st.session_state.err_ext = True

        if "err_int" not in st.session_state: st.session_state.err_int = False
        if "err_ext" not in st.session_state: st.session_state.err_ext = False

        st.markdown("###### Número de varales")
        val_act_var = st.session_state.get("num_var", 4)
        paso_var = 1 if val_act_var % 2 != 0 else 2
        num_varales = st.number_input("Número de varales", min_value=2, max_value=16, value=4, step=paso_var, key="num_var", label_visibility="collapsed")

        if num_varales == 2: sug_i, sug_e = 0, total_t // 2
        else:
            ideal_i = total_t / (num_varales + 2)
            sug_i = max(2, int(round(ideal_i / 2) * 2))
            sug_e = int((total_t - (num_varales - 2) * sug_i) / 2)

        if "last_max" not in st.session_state or st.session_state.last_max != total_t or \
           "last_var" not in st.session_state or st.session_state.last_var != num_varales:
            st.session_state.cap_ext_edit = max(2, sug_e) # Evitamos negativos por seguridad
            st.session_state.cap_int_edit = max(2, sug_i)
            st.session_state.err_int = False
            st.session_state.err_ext = False
            st.session_state.last_max = total_t
            st.session_state.last_var = num_varales

        if num_varales >= 4:
            st.write("")
            st.markdown("###### Capacidad de varales exteriores")
            capacidad_ext = st.number_input("Capacidad de varales exteriores", min_value=2, key="cap_ext_edit", step=1, on_change=actualizar_desde_exterior, label_visibility="collapsed")
            if st.session_state.err_ext: st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
            
            st.write("")
            st.markdown("###### Capacidad de varales interiores")
            val_act_int = st.session_state.get("cap_int_edit", 2)
            paso_int = 1 if val_act_int % 2 != 0 else 2
            capacidad_int = st.number_input("Capacidad de varales interiores", min_value=2, key="cap_int_edit", step=paso_int, on_change=actualizar_desde_interior, label_visibility="collapsed")
            if st.session_state.err_int: st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
            elif capacidad_int % 2 != 0: st.error("✖ &nbsp;&nbsp;&nbsp;La capacidad interior debe ser par.")
        else:
            capacidad_ext = total_t // 2
            capacidad_int = 0

        # Validación final cruzada
        suma_calculada = (num_varales - 2) * capacidad_int + 2 * capacidad_ext if num_varales > 2 else 2 * capacidad_ext
        cuadra_total = (suma_calculada == total_t)
        
        if not cuadra_total and not st.session_state.err_int and not st.session_state.err_ext:
            st.error(f"✖ &nbsp;&nbsp;&nbsp;La capacidad total configurada ({suma_calculada}) no coincide con los Titulares del Excel ({total_t}).")

        parametros_validos = (total_t % 2 == 0) and (num_varales % 2 == 0) and \
                             (capacidad_int % 2 == 0 or num_varales == 2) and \
                             not st.session_state.err_int and not st.session_state.err_ext and \
                             cuadra_total

        if parametros_validos:
            varales_por_lado = num_varales // 2
            for lado in ["Izquierdo", "Derecho"]:
                for i in range(varales_por_lado):
                    if i == 0: varales_config.append({"Nombre": f"Varal {lado} Exterior", "Lado": lado, "Capacidad": capacidad_ext, "Tipo": "Exterior"})
                    else:
                        sufijo = "" if varales_por_lado == 2 else f" {i}"
                        varales_config.append({"Nombre": f"Varal {lado} Interior{sufijo}", "Lado": lado, "Capacidad": capacidad_int, "Tipo": "Interior"})

        st.write("")
    if st.session_state.df_titulares is None:
        pass # Mensaje manejado arriba
    elif not parametros_validos: 
        st.warning("⚠ &nbsp;&nbsp;&nbsp;Corrige los parámetros para poder analizar el archivo.")
    else: 
        st.success("✔ &nbsp;&nbsp;&nbsp;Parámetros correctos. Pulsa 'Analizar'.")

    btn_analizar = st.button("Analizar", use_container_width=True, disabled=(not parametros_validos or archivo is None))


# ==========================================
# COLUMNA 3: RESULTADOS Y DESCARGA
# ==========================================
with col3:
    st.markdown("#### Resultados")
    st.write("")

    mensaje_estado = st.empty()

    if st.session_state.analisis_completado and archivo is not None:
        mensaje_estado.success("✔ &nbsp;&nbsp;&nbsp;¡Tallaje completado! Ya puedes descargar el cuadrante.")
    else:
        mensaje_estado.info("ⓘ &nbsp;&nbsp;&nbsp;Pulsa 'Analizar' para generar el cuadrante.")

    progress_text = st.empty()
    progress_bar = st.empty()
    boton_placeholder = st.empty()
    boton_descarga_bloqueado = not st.session_state.analisis_completado or archivo is None
    boton_placeholder.download_button(
        label="Descargar",
        data=st.session_state.excel_buffer if st.session_state.analisis_completado else b"",
        file_name=f"cuadrante_costaleros_{año_actual}.xlsx",
        use_container_width=True,
        type="primary",
        disabled=boton_descarga_bloqueado,
        key="btn_descarga_inicial"
    )

    if btn_analizar:

        mensaje_estado.empty()  # Limpiar mensaje anterior

        progress_text.write("Generando cuadrante...\n\nProgreso &nbsp; ⟶ &nbsp; **0%**")
        progress_bar.progress(0)
        _start_time = time.time()
        TOTAL_UNITS = 11  # 1 principal + 10 cambio

        def _update_progress(unit, sub=0, sub_total=1):
            pct = min((unit + sub / sub_total) / TOTAL_UNITS, 1.0)
            elapsed = time.time() - _start_time
            if pct > 0.01:
                remaining = elapsed / pct * (1 - pct)
                mins, secs = divmod(int(remaining), 60)
                time_text = f"{mins} min {secs} s" if mins else f"{secs} s"
                progress_text.write(f"Generando cuadrante...\n\nProgreso &nbsp; ⟶ &nbsp; **{int(pct*100)}%** &nbsp;|&nbsp; **~ {time_text}** restantes")
            else:
                progress_text.write("Generando cuadrante...\n\nProgreso &nbsp; ⟶ &nbsp; **0%**")
            progress_bar.progress(pct)

        if True:
            df = st.session_state.df_titulares.copy()

            # ── 1. LIMPIEZA DE DATOS ──────────────────────────────────────────────
            for col in ['Altura Hombro Izquierdo (cm)', 'Altura Hombro Derecho (cm)']:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', '.'), errors='coerce'
                )
            df['Altura Hombro Izquierdo (cm)'] = (
                df['Altura Hombro Izquierdo (cm)']
                .fillna(df['Altura Hombro Derecho (cm)']).fillna(0)
            )
            df['Altura Hombro Derecho (cm)'] = (
                df['Altura Hombro Derecho (cm)']
                .fillna(df['Altura Hombro Izquierdo (cm)']).fillna(0)
            )
            df['Altura_Media'] = (
                df['Altura Hombro Izquierdo (cm)'] + df['Altura Hombro Derecho (cm)']
            ) / 2.0
            pool = df.sort_values('Altura_Media', ascending=False).to_dict('records')

            max_filas = max(capacidad_ext, capacidad_int) if num_varales > 2 else capacidad_ext

            # ── 2. EJECUTAR ALGORITMO PRINCIPAL ──────────────────────────────────
            def _cb_principal(restart, total):
                _update_progress(0, restart, total)

            result = run_assignment_algorithm(pool, varales_config, capacidad_ext, capacidad_int, num_varales, max_filas,
                                              progress_callback=_cb_principal)
            asignaciones_por_varal, grid, all_positions, varal_filas, fila_positions, fila_activa, h_carga, varal_lado = result
            _update_progress(1)

            # ── 3. CUADRANTE DE CAMBIO DE HOMBRO ────────────────────────────────
            # Contar costaleros fijos por lado
            SD_count = 0  # 'Solo Derecho' -> fijados en lado Izquierdo
            SI_count = 0  # 'Solo Izquierdo' -> fijados en lado Derecho
            costaleros_con_lado = []  # (costalero_dict, lado_actual)

            for vname, idx, lado, tipo in all_positions:
                c = grid[vname][idx]
                pref = c['Preferencia de Hombro']
                if pref == 'Solo Derecho':
                    SD_count += 1
                elif pref == 'Solo Izquierdo':
                    SI_count += 1
                costaleros_con_lado.append((c, lado))

            N_por_lado = len(all_positions) // 2
            surplus = abs(SD_count - SI_count)

            # Seleccionar los flexibles que deben quedarse (surplus)
            stayers = set()
            if surplus > 0:
                # El surplus está en el lado con MENOS fijos (más flexibles que huecos disponibles)
                if SD_count > SI_count:
                    lado_surplus = 'Derecho'
                else:
                    lado_surplus = 'Izquierdo'

                # Candidatos a quedarse: flexibles en el lado surplus
                # Prioridad: 0=pref blanda satisfecha, 1=indiferente, 2=pref blanda no satisfecha
                candidatos_stay = []
                for c, lado in costaleros_con_lado:
                    if lado == lado_surplus and c['Preferencia de Hombro'] not in ('Solo Izquierdo', 'Solo Derecho'):
                        diff_hombros = abs(c['Altura Hombro Izquierdo (cm)'] - c['Altura Hombro Derecho (cm)'])
                        pref = c['Preferencia de Hombro']
                        pref_ok = (pref == 'Derecho' and lado == 'Izquierdo') or (pref == 'Izquierdo' and lado == 'Derecho')
                        if pref_ok:
                            prioridad = 0
                        elif pref == 'Indiferente':
                            prioridad = 1
                        else:
                            prioridad = 2
                        candidatos_stay.append((prioridad, diff_hombros, id(c), c))

                candidatos_stay.sort(key=lambda x: (x[0], x[1]))
                for i in range(min(surplus, len(candidatos_stay))):
                    stayers.add(id(candidatos_stay[i][3]))

            # Construir pool de cambio de hombro
            cambio_pool = []
            for c_original, lado_actual in costaleros_con_lado:
                c_new = dict(c_original)
                pref = c_original['Preferencia de Hombro']

                if pref in ('Solo Izquierdo', 'Solo Derecho'):
                    pass  # mantener preferencia fija
                elif id(c_original) in stayers:
                    # Preferir el lado actual (soft, permite excepciones para cross-row)
                    if lado_actual == 'Izquierdo':
                        c_new['Preferencia de Hombro'] = 'Derecho'   # prefiere lado Izquierdo
                    else:
                        c_new['Preferencia de Hombro'] = 'Izquierdo'  # prefiere lado Derecho
                else:
                    # Cambiar de lado: preferir el lado contrario (soft)
                    if lado_actual == 'Izquierdo':
                        c_new['Preferencia de Hombro'] = 'Izquierdo'  # prefiere lado Derecho
                    else:
                        c_new['Preferencia de Hombro'] = 'Derecho'    # prefiere lado Izquierdo

                cambio_pool.append(c_new)

            cambio_pool.sort(key=lambda x: x['Altura_Media'], reverse=True)

            # Ejecutar algoritmo para cambio de hombro — selección lexicográfica
            # Prioridades: 1) 0 cross-row  2) max cambios hombro  3) min ext/int  4) menor diff pares
            cambio_weights = {
                'W_CROSS': 10000.0,
                'W_PREF': 100.0,
                'W_EXT': 5.0,
                'W_ROW': 5.0,
                'W_PAIR': 2.0,
                'W_GRAD': 3.0,
            }

            # Mapa lado y preferencia original de cada costalero
            lado_original = {}
            pref_original = {}
            for c, lado in costaleros_con_lado:
                lado_original[c['Nombre']] = lado
                pref_original[c['Nombre']] = c['Preferencia de Hombro']

            N_CANDIDATES = 10
            best_candidate = None
            best_score = None  # (cross_row_violations, -cambios_hombro, ext_int_viol, pair_diff)

            for seed_offset in range(N_CANDIDATES):
                def _cb_cambio(restart, total, _so=seed_offset):
                    _update_progress(1 + _so, restart, total)

                cambio_result = run_assignment_algorithm(
                    cambio_pool, varales_config, capacidad_ext, capacidad_int,
                    num_varales, max_filas,
                    custom_weights=cambio_weights, relax_cross_row=True,
                    seed_offset=seed_offset, progress_callback=_cb_cambio
                )
                cand_asig = cambio_result[0]
                cand_grid = cambio_result[1]
                cand_fpos = cambio_result[4]
                cand_hcarga = cambio_result[6]
                cand_filas_ord = sorted(cand_fpos.keys())

                # ── Safety net: garantizar 0 cross-row ──

                def _crc(g):
                    total = 0.0
                    for ri in range(len(cand_filas_ord) - 1):
                        fc, fn = cand_filas_ord[ri], cand_filas_ord[ri + 1]
                        mn = min(cand_hcarga(g[vn][idx], lado) for vn, idx, lado in cand_fpos[fc])
                        mx = max(cand_hcarga(g[vn][idx], lado) for vn, idx, lado in cand_fpos[fn])
                        if mx > mn:
                            total += (mx - mn)
                    return total

                all_cpos = []
                for f in cand_filas_ord:
                    for vn, idx, lado in cand_fpos[f]:
                        all_cpos.append((vn, idx, lado, f))

                # Fase 1: greedy swaps same-side
                for _ in range(500):
                    cc = _crc(cand_grid)
                    if cc < 0.001:
                        break
                    bs, bc = None, cc
                    for i in range(len(all_cpos)):
                        for j in range(i + 1, len(all_cpos)):
                            vn_a, idx_a, la, fa = all_cpos[i]
                            vn_b, idx_b, lb, fb = all_cpos[j]
                            if fa == fb:
                                continue
                            ca, cb = cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b]
                            if not hard_ok(cb, la) or not hard_ok(ca, lb):
                                continue
                            cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b] = cb, ca
                            nc = _crc(cand_grid)
                            if nc < bc - 0.001:
                                bc, bs = nc, (vn_a, idx_a, vn_b, idx_b)
                            cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b] = ca, cb
                    if not bs:
                        break
                    va, ia, vb, ib = bs
                    cand_grid[va][ia], cand_grid[vb][ib] = cand_grid[vb][ib], cand_grid[va][ia]

                # Fase 2: sort por varal
                if _crc(cand_grid) > 0.001:
                    for v in varales_config:
                        vn, lado = v['Nombre'], v['Lado']
                        cand_grid[vn].sort(key=lambda c: cand_hcarga(c, lado), reverse=True)

                # Fase 3: any swap
                for _ in range(500):
                    cc = _crc(cand_grid)
                    if cc < 0.001:
                        break
                    bs, bc = None, cc
                    for i in range(len(all_cpos)):
                        for j in range(i + 1, len(all_cpos)):
                            vn_a, idx_a, la, fa = all_cpos[i]
                            vn_b, idx_b, lb, fb = all_cpos[j]
                            if fa == fb:
                                continue
                            ca, cb = cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b]
                            if not hard_ok(cb, la) or not hard_ok(ca, lb):
                                continue
                            cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b] = cb, ca
                            nc = _crc(cand_grid)
                            if nc < bc - 0.001:
                                bc, bs = nc, (vn_a, idx_a, vn_b, idx_b)
                            cand_grid[vn_a][idx_a], cand_grid[vn_b][idx_b] = ca, cb
                    if not bs:
                        break
                    va, ia, vb, ib = bs
                    cand_grid[va][ia], cand_grid[vb][ib] = cand_grid[vb][ib], cand_grid[va][ia]

                # Fase 4: rotaciones 3 vías
                for _ in range(200):
                    cb4 = _crc(cand_grid)
                    if cb4 < 0.001:
                        break
                    imp = False
                    for i in range(len(all_cpos)):
                        if imp: break
                        for j in range(len(all_cpos)):
                            if imp: break
                            if i == j: continue
                            for k in range(len(all_cpos)):
                                if k == i or k == j: continue
                                vi, ii, li, _ = all_cpos[i]
                                vj, ij, lj, _ = all_cpos[j]
                                vk, ik, lk, _ = all_cpos[k]
                                ci_v, cj_v, ck_v = cand_grid[vi][ii], cand_grid[vj][ij], cand_grid[vk][ik]
                                if not hard_ok(cj_v, li) or not hard_ok(ck_v, lj) or not hard_ok(ci_v, lk):
                                    continue
                                cand_grid[vi][ii], cand_grid[vj][ij], cand_grid[vk][ik] = cj_v, ck_v, ci_v
                                if _crc(cand_grid) < cb4 - 0.001:
                                    imp = True
                                    break
                                cand_grid[vi][ii], cand_grid[vj][ij], cand_grid[vk][ik] = ci_v, cj_v, ck_v
                    if not imp:
                        break

                # Fase 5: redistribución global + bubble sort
                if _crc(cand_grid) > 0.001:
                    all_c = []
                    for f in cand_filas_ord:
                        for vn, idx, lado in cand_fpos[f]:
                            all_c.append(cand_grid[vn][idx])
                    all_c.sort(key=lambda c: c['Altura_Media'], reverse=True)
                    ci_p = 0
                    for f in cand_filas_ord:
                        positions = cand_fpos[f]
                        n = len(positions)
                        row_c = all_c[ci_p:ci_p + n]
                        ci_p += n
                        used = [False] * n
                        for vn, idx, lado in positions:
                            bj, bh = -1, -1.0
                            for j in range(n):
                                if used[j]: continue
                                if not hard_ok(row_c[j], lado): continue
                                h = cand_hcarga(row_c[j], lado)
                                if h > bh:
                                    bh, bj = h, j
                            if bj == -1:
                                for j in range(n):
                                    if used[j]: continue
                                    h = cand_hcarga(row_c[j], lado)
                                    if h > bh:
                                        bh, bj = h, j
                            cand_grid[vn][idx] = row_c[bj]
                            used[bj] = True
                    for _ in range(1000):
                        if _crc(cand_grid) < 0.001: break
                        sw = False
                        for ri in range(len(cand_filas_ord) - 1):
                            fc, fn = cand_filas_ord[ri], cand_filas_ord[ri + 1]
                            mn = min(cand_hcarga(cand_grid[vn][idx], lado) for vn, idx, lado in cand_fpos[fc])
                            mx = max(cand_hcarga(cand_grid[vn][idx], lado) for vn, idx, lado in cand_fpos[fn])
                            if mx <= mn: continue
                            ps = min(cand_fpos[fc], key=lambda t: cand_hcarga(cand_grid[t[0]][t[1]], t[2]))
                            pt = max(cand_fpos[fn], key=lambda t: cand_hcarga(cand_grid[t[0]][t[1]], t[2]))
                            va, ia, la = ps
                            vb, ib, lb = pt
                            if hard_ok(cand_grid[vb][ib], la) and hard_ok(cand_grid[va][ia], lb):
                                cand_grid[va][ia], cand_grid[vb][ib] = cand_grid[vb][ib], cand_grid[va][ia]
                                sw = True
                        if not sw: break

                # ── Evaluar este candidato según prioridades lexicográficas ──

                # 1) Cross-row violations
                cross_viol = 0
                for ri in range(len(cand_filas_ord) - 1):
                    fc, fn = cand_filas_ord[ri], cand_filas_ord[ri + 1]
                    hc = [cand_hcarga(cand_grid[vn][idx], lado) for vn, idx, lado in cand_fpos[fc]]
                    hn = [cand_hcarga(cand_grid[vn][idx], lado) for vn, idx, lado in cand_fpos[fn]]
                    if max(hn) > min(hc):
                        cross_viol += 1

                # 2) Cambios de hombro (más = mejor → negamos para minimizar)
                n_cambios = 0
                for v in varales_config:
                    for c in cand_grid[v['Nombre']]:
                        if v['Lado'] != lado_original.get(c['Nombre'], v['Lado']):
                            n_cambios += 1

                # 2b) Calidad de no-cambios: penalizar si los que no cambian tenían pref blanda no satisfecha
                no_cambio_penalty = 0
                for v in varales_config:
                    for c in cand_grid[v['Nombre']]:
                        if v['Lado'] == lado_original.get(c['Nombre'], v['Lado']):
                            pref = pref_original.get(c['Nombre'], 'Indiferente')
                            lado_princ = lado_original.get(c['Nombre'])
                            pref_ok = (pref == 'Derecho' and lado_princ == 'Izquierdo') or (pref == 'Izquierdo' and lado_princ == 'Derecho')
                            if pref in ('Derecho', 'Izquierdo') and not pref_ok:
                                no_cambio_penalty += 2
                            elif pref == 'Indiferente':
                                no_cambio_penalty += 1

                # 3) Violaciones ext/int
                ext_viol = 0
                for f, pos_list in cand_fpos.items():
                    for side in ('Izquierdo', 'Derecho'):
                        ext_h = [cand_hcarga(cand_grid[vn][ri], lado) for vn, ri, lado in pos_list
                                 if lado == side and any(v['Nombre'] == vn and v['Tipo'] == 'Exterior' for v in varales_config)]
                        int_h = [cand_hcarga(cand_grid[vn][ri], lado) for vn, ri, lado in pos_list
                                 if lado == side and any(v['Nombre'] == vn and v['Tipo'] == 'Interior' for v in varales_config)]
                        if ext_h and int_h and max(int_h) > min(ext_h):
                            ext_viol += 1

                # 4) Diferencia media de pares
                pair_diffs = []
                for v in varales_config:
                    if v['Lado'] != 'Izquierdo':
                        continue
                    vl, tipo = v['Nombre'], v['Tipo']
                    vr_name = [vr['Nombre'] for vr in varales_config if vr['Lado'] == 'Derecho' and vr['Tipo'] == tipo][0]
                    for idx in range(min(len(cand_grid[vl]), len(cand_grid[vr_name]))):
                        hl = cand_hcarga(cand_grid[vl][idx], 'Izquierdo')
                        hr = cand_hcarga(cand_grid[vr_name][idx], 'Derecho')
                        pair_diffs.append(abs(hl - hr))
                avg_pair = sum(pair_diffs) / len(pair_diffs) if pair_diffs else 0

                score = (cross_viol, -n_cambios, no_cambio_penalty, ext_viol, avg_pair)

                if best_score is None or score < best_score:
                    best_score = score
                    best_candidate = {
                        'asignaciones': {vn: list(cs) for vn, cs in cand_asig.items()},
                        'grid': {vn: list(cs) for vn, cs in cand_grid.items()},
                        'result': cambio_result,
                    }

                _update_progress(1 + seed_offset + 1)

            # Usar el mejor candidato
            cambio_asignaciones = best_candidate['asignaciones']
            cambio_grid = best_candidate['grid']
            cambio_result = best_candidate['result']
            cambio_fila_activa = cambio_result[5]
            cambio_h_carga = cambio_result[6]
            cambio_fila_positions = cambio_result[4]
            cambio_varal_filas = cambio_result[3]

            # ── Post-procesado: intentar swaps para reducir BLANDA NO SAT ──
            def _calc_score(g, fpos, hcfn):
                filas_ord = sorted(fpos.keys())
                cr = 0
                for ri in range(len(filas_ord) - 1):
                    fc, fn = filas_ord[ri], filas_ord[ri + 1]
                    hc = [hcfn(g[vn][idx], lado) for vn, idx, lado in fpos[fc]]
                    hn = [hcfn(g[vn][idx], lado) for vn, idx, lado in fpos[fn]]
                    if max(hn) > min(hc): cr += 1
                nc = 0
                for v in varales_config:
                    for c in g[v['Nombre']]:
                        if v['Lado'] != lado_original.get(c['Nombre'], v['Lado']):
                            nc += 1
                ncp = 0
                for v in varales_config:
                    for c in g[v['Nombre']]:
                        if v['Lado'] == lado_original.get(c['Nombre'], v['Lado']):
                            p = pref_original.get(c['Nombre'], 'Indiferente')
                            lp = lado_original.get(c['Nombre'])
                            pok = (p == 'Derecho' and lp == 'Izquierdo') or (p == 'Izquierdo' and lp == 'Derecho')
                            if p in ('Derecho', 'Izquierdo') and not pok: ncp += 2
                            elif p == 'Indiferente': ncp += 1
                ev = 0
                for f, pl in fpos.items():
                    for side in ('Izquierdo', 'Derecho'):
                        eh = [hcfn(g[vn][ri], lado) for vn, ri, lado in pl
                              if lado == side and any(v['Nombre'] == vn and v['Tipo'] == 'Exterior' for v in varales_config)]
                        ih = [hcfn(g[vn][ri], lado) for vn, ri, lado in pl
                              if lado == side and any(v['Nombre'] == vn and v['Tipo'] == 'Interior' for v in varales_config)]
                        if eh and ih and max(ih) > min(eh): ev += 1
                pd_list = []
                for v in varales_config:
                    if v['Lado'] != 'Izquierdo': continue
                    vl, tipo = v['Nombre'], v['Tipo']
                    vr = [vr['Nombre'] for vr in varales_config if vr['Lado'] == 'Derecho' and vr['Tipo'] == tipo][0]
                    for idx in range(min(len(g[vl]), len(g[vr]))):
                        pd_list.append(abs(hcfn(g[vl][idx], 'Izquierdo') - hcfn(g[vr][idx], 'Derecho')))
                ap = sum(pd_list) / len(pd_list) if pd_list else 0
                return (cr, -nc, ncp, ev, ap)

            pre_score = _calc_score(cambio_grid, cambio_fila_positions, cambio_h_carga)

            # Identificar no-cambios BLANDA NO SAT y candidatos a swap
            _no_cambia_bad = []
            _swap_targets = []
            for v in varales_config:
                for idx, c in enumerate(cambio_grid[v['Nombre']]):
                    nombre = c['Nombre']
                    lado_c = v['Lado']
                    if lado_c == lado_original.get(nombre, lado_c):
                        pref = pref_original[nombre]
                        lp = lado_original[nombre]
                        pok = (pref == 'Derecho' and lp == 'Izquierdo') or (pref == 'Izquierdo' and lp == 'Derecho')
                        if pref in ('Derecho', 'Izquierdo') and not pok:
                            _no_cambia_bad.append((v['Nombre'], idx, lado_c))
                        elif pref not in ('Solo Derecho', 'Solo Izquierdo'):
                            _swap_targets.append((v['Nombre'], idx, lado_c))
                    else:
                        pref = pref_original[nombre]
                        if pref not in ('Solo Derecho', 'Solo Izquierdo'):
                            _swap_targets.append((v['Nombre'], idx, lado_c))

            for bad_vn, bad_idx, bad_lado in _no_cambia_bad:
                for tgt_vn, tgt_idx, tgt_lado in _swap_targets:
                    if bad_lado == tgt_lado:
                        continue
                    cambio_grid[bad_vn][bad_idx], cambio_grid[tgt_vn][tgt_idx] = cambio_grid[tgt_vn][tgt_idx], cambio_grid[bad_vn][bad_idx]
                    new_score = _calc_score(cambio_grid, cambio_fila_positions, cambio_h_carga)
                    if new_score <= pre_score and new_score[4] - pre_score[4] <= 0.1:
                        pre_score = new_score
                        break  # Aceptar swap
                    cambio_grid[bad_vn][bad_idx], cambio_grid[tgt_vn][tgt_idx] = cambio_grid[tgt_vn][tgt_idx], cambio_grid[bad_vn][bad_idx]

            # Actualizar cambio_asignaciones desde cambio_grid
            for vn in cambio_asignaciones:
                cambio_asignaciones[vn] = cambio_grid[vn]

            # ── 7b. SUPLENTES — BUSCAR MEJOR TITULAR POR HOMBRO ────────────
            suplentes_resultado = []
            df_supl = st.session_state.df_suplentes
            if df_supl is not None and len(df_supl) > 0:
                df_supl = df_supl.copy()
                for col_s in ['Altura Hombro Izquierdo (cm)', 'Altura Hombro Derecho (cm)']:
                    df_supl[col_s] = pd.to_numeric(
                        df_supl[col_s].astype(str).str.replace(',', '.'), errors='coerce')
                df_supl['Altura Hombro Izquierdo (cm)'] = (
                    df_supl['Altura Hombro Izquierdo (cm)']
                    .fillna(df_supl['Altura Hombro Derecho (cm)']).fillna(0))
                df_supl['Altura Hombro Derecho (cm)'] = (
                    df_supl['Altura Hombro Derecho (cm)']
                    .fillna(df_supl['Altura Hombro Izquierdo (cm)']).fillna(0))

                titulares_por_lado = {'Izquierdo': [], 'Derecho': []}
                for varal in varales_config:
                    nombre_v, lado_v, tipo_v = varal['Nombre'], varal['Lado'], varal['Tipo']
                    filas_v = [f for f in range(1, max_filas + 1) if fila_activa(f, tipo_v)]
                    for idx_v, c in enumerate(asignaciones_por_varal[nombre_v]):
                        pos_str = f"{nombre_v} - Fila {filas_v[idx_v]}"
                        titulares_por_lado[lado_v].append((c, h_carga(c, lado_v), pos_str))

                for _, supl in df_supl.iterrows():
                    s = supl.to_dict()
                    pref = s.get('Preferencia de Hombro', 'Indiferente')

                    h_izq_s = str(s['Altura Hombro Izquierdo (cm)']).replace('.', ',')
                    h_der_s = str(s['Altura Hombro Derecho (cm)']).replace('.', ',')
                    
                    # Guardamos el nombre limpio para mostrar, y las alturas para el tooltip
                    nombre_display = s['Nombre']
                    if pref == 'Solo Izquierdo':
                        altura_msg = f"{h_izq_s} cm"
                        titulo_msg = "Altura:"
                    elif pref == 'Solo Derecho':
                        altura_msg = f"{h_der_s} cm"
                        titulo_msg = "Altura:"
                    else:
                        altura_msg = f"{h_izq_s} cm\n{h_der_s} cm"
                        titulo_msg = "Alturas:"

                    if pref == 'Solo Derecho':
                        lados_s = ['Izquierdo']
                    elif pref == 'Solo Izquierdo':
                        lados_s = ['Derecho']
                    else:
                        lados_s = ['Izquierdo', 'Derecho']

                    # Ahora guardamos tuplas: (Nombre_A_Mostrar, Mensaje_Altura)
                    fila_data = {
                        'Nombre': (nombre_display, titulo_msg, altura_msg),
                        'sust_izq': ('', ''), 'pos_izq': '',
                        'sust_der': ('', ''), 'pos_der': ''
                    }

                    for lado_s in lados_s:
                        h_supl = h_carga(s, lado_s)
                        mejor, mejor_diff, mejor_h, mejor_pos = None, float('inf'), 0, ''
                        for titular, h_tit, p_str in titulares_por_lado[lado_s]:
                            d = abs(h_supl - h_tit)
                            if d < mejor_diff:
                                mejor, mejor_diff, mejor_h, mejor_pos = titular, d, h_tit, p_str
                        if mejor:
                            h_str = str(mejor_h).replace('.', ',')
                            # Separamos el nombre del titular y su altura
                            nombre_tit = mejor['Nombre']
                            msg_tit = f"{h_str} cm"
                            
                            if lado_s == 'Izquierdo':
                                fila_data['sust_der'] = (nombre_tit, msg_tit)
                                fila_data['pos_der'] = mejor_pos
                            else:
                                fila_data['sust_izq'] = (nombre_tit, msg_tit)
                                fila_data['pos_izq'] = mejor_pos

                    suplentes_resultado.append(fila_data)

            # ── 7. GENERACIÓN DEL EXCEL ───────────────────────────────────────────
    
            resultado = []
            for varal in varales_config:
                nombre, lado, tipo = varal["Nombre"], varal["Lado"], varal["Tipo"]
                filas_validas = [f for f in range(1, max_filas + 1) if fila_activa(f, tipo)]
                for i, c in enumerate(asignaciones_por_varal[nombre]):
                    resultado.append({
                        'Varal'  : nombre,
                        'Fila'   : filas_validas[i],
                        'Nombre' : c['Nombre'],
                        'Altura' : str(h_carga(c, lado)).replace('.', ',')
                    })
    
            columnas_nombres = [v["Nombre"] for v in varales_config]
            izq_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Exterior"]
            izq_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Interior"][::-1]
            der_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho"   and v["Tipo"] == "Interior"]
            der_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho"   and v["Tipo"] == "Exterior"]
            orden_final = izq_e + izq_i + der_i + der_e
    
            df_res = pd.DataFrame("", index=range(max_filas), columns=orden_final)
            for r in resultado:
                df_res.at[r['Fila'] - 1, r['Varal']] = (r['Nombre'], r['Altura'])

            # Resultado cambio de hombro
            resultado_cambio = []
            for varal in varales_config:
                nombre, lado, tipo = varal["Nombre"], varal["Lado"], varal["Tipo"]
                filas_validas = [f for f in range(1, max_filas + 1) if cambio_fila_activa(f, tipo)]
                for i, c in enumerate(cambio_asignaciones[nombre]):
                    resultado_cambio.append({
                        'Varal'  : nombre,
                        'Fila'   : filas_validas[i],
                        'Nombre' : c['Nombre'],
                        'Altura' : str(cambio_h_carga(c, lado)).replace('.', ',')
                    })

            df_cambio = pd.DataFrame("", index=range(max_filas), columns=orden_final)
            for r in resultado_cambio:
                df_cambio.at[r['Fila'] - 1, r['Varal']] = (r['Nombre'], r['Altura'])
    
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet('Cuadrante Principal')
                
                # Formatos
                titulo_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF2B2B', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#8583FF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 13})
                cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                index_fmt = workbook.add_format({'bold': True, 'bg_color': '#8583FF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 13})
                
                # --- 1. TÍTULO PRINCIPAL (Fila 1) ---
                worksheet.set_row(0, 35)
                # Combinar desde la columna A hasta la última de los varales
                worksheet.merge_range(0, 0, 0, len(orden_final) - 1, "Cuadrante Principal", titulo_fmt)
                
                # --- 2. CABECERAS DE VARALES E ÍNDICES (Fila 2) ---
                worksheet.set_row(1, 30)
                for i, col in enumerate(df_res.columns):
                    worksheet.set_column(i, i, 40)
                    worksheet.write(1, i, col, header_fmt)
                
                # Cabecera de la columna de números de fila (a la derecha) -> Totalmente en blanco
                col_indices = len(orden_final)
                worksheet.set_column(col_indices, col_indices, 5)
                    
                # --- 3. DATOS, MENSAJES INVISIBLES Y NÚMEROS DE FILA (Empiezan en Fila 3) ---
                for f in range(len(df_res)):
                    row_idx = f + 2  # Índice 2 en Excel es la Fila 3
                    worksheet.set_row(row_idx, 25)
                    
                    for i, col in enumerate(df_res.columns):
                        val = df_res.iloc[f, i]
                        if val != "":
                            nombre, altura = val
                            # 1. Escribir solo el nombre (celda totalmente limpia)
                            worksheet.write(row_idx, i, nombre, cell_fmt)
                            
                            # 2. Añadir el mensaje de entrada (Sin marca roja, aparece al hacer clic)
                            worksheet.data_validation(row_idx, i, row_idx, i, {
                                'validate': 'any',
                                'input_title': 'Altura:',
                                'input_message': f'{altura} cm'
                            })
                        else:
                            worksheet.write_blank(row_idx, i, "", cell_fmt)
                    
                    # Escribir el número de fila a la derecha del todo
                    worksheet.write(row_idx, col_indices, f + 1, index_fmt)
                
                # --- 4. BLOQUE GRIS CENTRAL ---
                if num_varales >= 4:
                    centro_fmt = workbook.add_format({'bg_color': "#BEBEBE", 'border': 1})
                    m_ini = (capacidad_int // 2) + 2
                    m_fin =  max_filas - (capacidad_int // 2) + 1
                    if m_ini <= m_fin:
                        worksheet.merge_range(m_ini, 1, m_fin, len(orden_final) - 2, "", centro_fmt)

                # --- 5. HOJA DE SUPLENTES ---
                if suplentes_resultado:
                    ws_s = workbook.add_worksheet('Suplentes')
                    ws_s.set_row(0, 35)
                    ws_s.merge_range(0, 0, 0, 4, "Costaleros Suplentes", titulo_fmt)

                    ws_s.set_row(1, 30)
                    cabeceras_s = ["Nombre", "Sustituto Hombro Izquierdo", "Posición",
                                   "Sustituto Hombro Derecho", "Posición"]
                    anchos_s = [40, 40, 40, 40, 40]
                    for i_s, (cab, ancho) in enumerate(zip(cabeceras_s, anchos_s)):
                        ws_s.set_column(i_s, i_s, ancho)
                        ws_s.write(1, i_s, cab, header_fmt)

                    for i_s, d_s in enumerate(suplentes_resultado):
                        row_s = i_s + 2
                        ws_s.set_row(row_s, 25)
                        
                        # Col 0: Nombre Suplente
                        nom_supl, titulo_supl, msg_supl = d_s['Nombre']
                        ws_s.write(row_s, 0, nom_supl, cell_fmt)
                        ws_s.data_validation(row_s, 0, row_s, 0, {
                            'validate': 'any', 'input_title': titulo_supl, 'input_message': msg_supl
                        })

                        # Col 1: Sustituto H. Izq
                        nom_tit_i, msg_tit_i = d_s['sust_izq']
                        if nom_tit_i:
                            ws_s.write(row_s, 1, nom_tit_i, cell_fmt)
                            ws_s.data_validation(row_s, 1, row_s, 1, {
                                'validate': 'any', 'input_title': 'Altura:', 'input_message': msg_tit_i
                            })
                        else:
                            ws_s.write_blank(row_s, 1, "", cell_fmt)

                        # Col 2: Posición Izq
                        ws_s.write(row_s, 2, d_s['pos_izq'], cell_fmt)

                        # Col 3: Sustituto H. Der
                        nom_tit_d, msg_tit_d = d_s['sust_der']
                        if nom_tit_d:
                            ws_s.write(row_s, 3, nom_tit_d, cell_fmt)
                            ws_s.data_validation(row_s, 3, row_s, 3, {
                                'validate': 'any', 'input_title': 'Altura:', 'input_message': msg_tit_d
                            })
                        else:
                            ws_s.write_blank(row_s, 3, "", cell_fmt)

                        # Col 4: Posición Der
                        ws_s.write(row_s, 4, d_s['pos_der'], cell_fmt)

                # --- 6. HOJA DE CAMBIO DE HOMBRO ---
                ws_c = workbook.add_worksheet('Cuadrante Cambio Hombro')

                ws_c.set_row(0, 35)
                ws_c.merge_range(0, 0, 0, len(orden_final) - 1, "Cuadrante Cambio de Hombro", titulo_fmt)

                # Detectar costaleros que no cambian de hombro
                lado_principal = {}
                for r in resultado:
                    varal_lado_map = {v['Nombre']: v['Lado'] for v in varales_config}
                    lado_principal[r['Nombre']] = varal_lado_map[r['Varal']]
                lado_cambio = {}
                for r in resultado_cambio:
                    varal_lado_map = {v['Nombre']: v['Lado'] for v in varales_config}
                    lado_cambio[r['Nombre']] = varal_lado_map[r['Varal']]
                no_cambia = {n for n in lado_principal if lado_principal[n] == lado_cambio.get(n)}

                no_cambio_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12, 'bg_color': '#FFF9C4'})

                ws_c.set_row(1, 30)
                for i, col in enumerate(df_cambio.columns):
                    ws_c.set_column(i, i, 40)
                    ws_c.write(1, i, col, header_fmt)

                col_indices_c = len(orden_final)
                ws_c.set_column(col_indices_c, col_indices_c, 5)

                for f in range(len(df_cambio)):
                    row_idx = f + 2
                    ws_c.set_row(row_idx, 25)

                    for i, col in enumerate(df_cambio.columns):
                        val = df_cambio.iloc[f, i]
                        if val != "":
                            nombre, altura = val
                            fmt = no_cambio_fmt if nombre in no_cambia else cell_fmt
                            ws_c.write(row_idx, i, nombre, fmt)
                            ws_c.data_validation(row_idx, i, row_idx, i, {
                                'validate': 'any',
                                'input_title': 'Altura:',
                                'input_message': f'{altura} cm'
                            })
                        else:
                            ws_c.write_blank(row_idx, i, "", cell_fmt)

                    ws_c.write(row_idx, col_indices_c, f + 1, index_fmt)

                if num_varales >= 4:
                    centro_fmt_c = workbook.add_format({'bg_color': "#BEBEBE", 'border': 1})
                    m_ini = (capacidad_int // 2) + 2
                    m_fin =  max_filas - (capacidad_int // 2) + 1
                    if m_ini <= m_fin:
                        ws_c.merge_range(m_ini, 1, m_fin, len(orden_final) - 2, "", centro_fmt_c)

                # Leyenda (con columna de espacio)
                leyenda_col = col_indices_c + 2  # +2 para dejar una columna de separación
                ws_c.set_column(leyenda_col, leyenda_col, 5)
                ws_c.set_column(leyenda_col + 1, leyenda_col + 1, 30)
                leyenda_color_fmt = workbook.add_format({'bg_color': '#FFF9C4', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                leyenda_text_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                ws_c.merge_range(1, leyenda_col, 1, leyenda_col + 1, "Leyenda", header_fmt)
                leyenda_white_fmt = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                ws_c.write(2, leyenda_col, "", leyenda_white_fmt)
                ws_c.write(2, leyenda_col + 1, "SÍ cambia de hombro", leyenda_text_fmt)
                ws_c.write(3, leyenda_col, "", leyenda_color_fmt)
                ws_c.write(3, leyenda_col + 1, "NO cambia de hombro", leyenda_text_fmt)

            st.session_state.excel_buffer       = output.getvalue()
            st.session_state.num_asignados      = len(resultado)
            st.session_state.analisis_completado = True

        progress_text.empty()
        progress_bar.empty()
        mensaje_estado.success("✔ &nbsp;&nbsp;&nbsp;¡Tallaje completado! Ya puedes descargar el cuadrante.")

        boton_placeholder.empty()
        boton_placeholder.download_button(
            label="Descargar",
            data=st.session_state.excel_buffer if st.session_state.analisis_completado else b"",
            file_name=f"cuadrante_costaleros_{año_actual}.xlsx",
            use_container_width=True,
            type="primary",
            disabled=False,
            key="btn_descarga_listo"
        )