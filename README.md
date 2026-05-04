# Production Planning System — Plastic Injection Manufacturing

> **Projeto I 2024/2025 · ISEP · MEGI / MEGCA**  
> Authors: Diogo Morgado Rainho · Filipa Sousa Neves  
> Supervisor: Professor Manuel Pereira Lopes

---

## Overview

This project implements a **greedy EDD-based (Earliest Due Date) production scheduling algorithm** in Python, developed for a plastic injection manufacturing company (*Nomoldes*). The goal is to optimise the assignment of work orders (OFs — *Ordens de Fabrico*) to machines over a 7-day planning horizon, minimising lateness while respecting real operational constraints.

The system evaluates **6 configurable scenarios** that vary the number of operators, SMED teams, and setup times, providing a comparative sensitivity analysis of how each parameter affects overall production efficiency.

---

## Features

- **EDD scheduling** with secondary tool-grouping to reduce unnecessary setup changes
- **Operator pool management** — tasks are deferred until sufficient operators are free
- **SMED team capacity** — concurrent setups are capped based on available teams
- **Primary + alternative machine assignment** — evaluated in priority order (primary → alt1 → alt2 → alt3)
- **Setup time logic** — charged only when the tool changes; waived on a machine's first use
- **Multi-scenario analysis** — 6 parameter configurations run and compared automatically
- **Gantt chart generation** with 7-day horizon marker
- **Excel output** — results written back to the corresponding sheet per scenario

---

## Algorithm Summary

```
Sort work orders by (due_date ASC, tool)
For each work order:
    Wait until enough operators are free
    Wait until a SMED team slot is available
    Evaluate all candidate machines (primary + up to 3 alternatives):
        compute setup time (0 if first use or same tool, else base_setup_time)
        compute finish time = max(now, machine_free_time) + setup + processing_time
    Assign to machine with earliest finish (tie-break: setup time, then slack)
    Update machine state, operator pool, SMED slots
    Record lateness / earliness / setup KPIs
```

---

## Scenarios

| Scenario | Operators | SMED Teams | Setup Time (min) |
|----------|-----------|------------|-----------------|
| 1 (baseline) | 35 | 1 | 60 |
| 2 | 30 | 1 | 60 |
| 3 | 35 | 2 | 60 |
| 4 | 35 | 1 | 45 |
| 5 | 40 | 1 | 60 |
| 6 | 35 | 3 | 60 |

---

## Results Summary

| | Sc. 1 | Sc. 2 | Sc. 3 | Sc. 4 | Sc. 5 | Sc. 6 |
|---|---|---|---|---|---|---|
| **Total Lateness (min)** | 51 020 | 117 683 | 30 397 | 31 667 | **30 163** | 30 617 |
| **Total Setup Time (min)** | 16 620 | 16 980 | 16 740 | **12 555** | 16 980 | 16 680 |
| **Avg Earliness (days)** | 1.05 | 0.87 | 1.16 | 1.17 | **1.20** | 1.16 |
| **Orders Processed** | 384 | 384 | 384 | 384 | 384 | 384 |

**Scenario 5** (40 operators, 1 SMED team) achieved the best overall performance — lowest lateness and highest average earliness — suggesting that operator availability is the primary bottleneck in this system.

---

## Project Structure

```
.
├── solution.py          # Main scheduling script
├── data/
│   ├── PRJT1.xlsx       # ⚠ Input dataset (not included — see below)
│   └── Results.xlsx     # Output file with one sheet per scenario
└── README.md
```

---

## ⚠ Data Availability

The input dataset (`PRJT1.xlsx`) contains confidential production data from a real manufacturing company and **cannot be shared for privacy reasons**. The output results file is similarly excluded.

To run the scheduler with your own data, your Excel file must contain the following columns:

| Column | Description |
|--------|-------------|
| `OF` | Work order identifier |
| `Due_date` | Due date in days (planning horizon: 7 days) |
| `Proc_time(min)` | Processing time in minutes |
| `Machine` | Primary machine code |
| `Tool` | Tool/mould identifier |
| `MachAlter1` | Alternative machine 1 (optional) |
| `MachAlter2` | Alternative machine 2 (optional) |
| `MachAlter3` | Alternative machine 3 (optional) |
| `Nb_Operator` | Number of operators required |

---

## Setup & Usage

### Requirements

```bash
pip install pandas matplotlib openpyxl
```

### Configuration

Edit the paths and scenario parameters at the top of `solution.py`:

```python
INPUT_FILE_PATH  = os.path.join(os.path.dirname(__file__), "data", "PRJT1.xlsx")
OUTPUT_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "Results.xlsx")
```

### Run

```bash
python solution.py
```

The script will:
1. Load and preprocess the work orders
2. Run all 6 scenarios sequentially
3. Print KPIs for each scenario to the console
4. Display a Gantt chart per scenario
5. Write results to the output Excel file

---

## Limitations & Future Work

- The current approach is a **greedy heuristic** — it does not guarantee a globally optimal solution
- No backtracking: once a task is assigned, it is not reconsidered
- For better solutions, **meta-heuristic methods** (e.g. Genetic Algorithms, Simulated Annealing, Tabu Search) could be applied
- The algorithm assumes continuous 24/7 machine availability; shift patterns or maintenance windows are not modelled

---

## References

- Bastos, J. (2024). *Algoritmos de suporte ao Escalonamento*. ISEP.
- Bastos, J., & Santos, A. (2024). *Approximate Techniques*. ISEP.
- Python Software Foundation. (2025). *Python 3 Documentation*. https://docs.python.org/3/
