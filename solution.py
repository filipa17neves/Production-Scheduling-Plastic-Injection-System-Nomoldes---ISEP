"""
Production Planning System for Plastic Injection Manufacturing
==============================================================
Implements an EDD-based (Earliest Due Date) scheduling algorithm with:
- Operator availability constraints
- SMED (Single-Minute Exchange of Die) team capacity management
- Primary and alternative machine assignment
- Multi-scenario analysis for operational parameter sensitivity

Authors: Diogo Morgado Rainho, Filipa Sousa Neves
Institution: ISEP – Instituto Superior de Engenharia do Porto
Course: Projeto I 2024/2025 – MEGI/MEGCA
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl


# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

# Path to the input dataset — update this to your local Excel file path.
# The dataset is not included in this repository for privacy reasons.
INPUT_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "PRJT1.xlsx")

# Path where scheduled results will be written (one sheet per scenario).
OUTPUT_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "Results.xlsx")

# Columns required in the input sheet
REQUIRED_COLUMNS = [
    "OF", "Due_date", "Proc_time(min)", "Machine", "Tool",
    "MachAlter1", "MachAlter2", "MachAlter3", "Nb_Operator"
]

# Minutes in a full day — used to convert due dates (given in days) to minutes
MINUTES_PER_DAY = 1440

# Scenario definitions: each scenario varies operators, SMED teams, and setup time.
# Scenario 1 is the baseline; the others explore parameter sensitivity.
SCENARIOS = {
    "Cenário 1": {"n_operators": 35, "n_smed_teams": 1, "base_setup_time": 60},
    "Cenário 2": {"n_operators": 30, "n_smed_teams": 1, "base_setup_time": 60},
    "Cenário 3": {"n_operators": 35, "n_smed_teams": 2, "base_setup_time": 60},
    "Cenário 4": {"n_operators": 35, "n_smed_teams": 1, "base_setup_time": 45},
    "Cenário 5": {"n_operators": 40, "n_smed_teams": 1, "base_setup_time": 60},
    "Cenário 6": {"n_operators": 35, "n_smed_teams": 3, "base_setup_time": 60},
}


# ---------------------------------------------------------------------------
# DATA LOADING & PREPROCESSING
# ---------------------------------------------------------------------------

def load_and_prepare_data(file_path: str) -> tuple[pd.DataFrame, int, int]:
    """
    Load the production orders from Excel and apply preprocessing steps:
      1. Validate required columns exist.
      2. Convert due dates from days to minutes.
      3. Merge duplicate work orders (same OF reference) by summing processing times.
      4. Absorb orders with due_date == 8 into the preceding order (plant-specific rule:
         a tool must remain on the machine for at least 240 min, so day-8 orders are
         appended to the last day-7 order rather than scheduled independently).

    Returns:
        data_final     – cleaned DataFrame ready for scheduling
        n_grouped      – number of work orders collapsed during deduplication
        n_day8_orders  – number of day-8 orders absorbed into prior orders
    """
    try:
        raw = pd.read_excel(file_path, sheet_name=None)
        sheet_name = list(raw.keys())[0]
        data = raw[sheet_name]
    except FileNotFoundError:
        raise FileNotFoundError(
            f"Input file not found at '{file_path}'.\n"
            "Please update INPUT_FILE_PATH in the configuration section."
        )

    # Validate schema
    for col in REQUIRED_COLUMNS:
        if col not in data.columns:
            raise KeyError(f"Required column '{col}' is missing from the Excel sheet.")

    # Convert due date from days → minutes for uniform time arithmetic
    data["Due_date_min"] = data["Due_date"] * MINUTES_PER_DAY

    # --- Step 1: Merge duplicate work orders ---
    # Multiple rows with the same OF reference represent split quantities of the
    # same job; we consolidate them into a single order with summed processing time.
    data_grouped = (
        data.groupby("OF")
        .agg({
            "Proc_time(min)": "sum",
            "Due_date_min":   "first",
            "Machine":        "first",
            "Tool":           "first",
            "MachAlter1":     "first",
            "MachAlter2":     "first",
            "MachAlter3":     "first",
            "Nb_Operator":    "first",
        })
        .reset_index()
    )
    n_grouped = len(data) - len(data_grouped)

    # --- Step 2: Absorb day-8 orders ---
    # Due_date == 8 exceeds the 7-day planning horizon. These arise because the mould
    # must stay in the machine for a minimum run; we merge their processing time into
    # the immediately preceding order so the schedule stays within bounds.
    n_day8_orders = int((data["Due_date"] == 8).sum())

    data_sorted = data_grouped.sort_values("Due_date_min").reset_index(drop=True)
    consolidated = []
    previous = None

    for _, row in data_sorted.iterrows():
        if row["Due_date_min"] == 8 * MINUTES_PER_DAY and previous is not None:
            previous["Proc_time(min)"] += row["Proc_time(min)"]
        else:
            previous = row.copy()
            consolidated.append(previous)

    data_final = pd.DataFrame(consolidated)
    return data_final, n_grouped, n_day8_orders


# ---------------------------------------------------------------------------
# SCHEDULING ENGINE
# ---------------------------------------------------------------------------

class ProductionScheduler:
    """
    Greedy EDD scheduler with resource constraints.

    The scheduler processes work orders sorted by (due_date, tool) — primary EDD
    ordering, with secondary tool-grouping to reduce unnecessary setup changes.
    For each order it selects the best available machine by minimising completion
    time, breaking ties on setup time and then slack.

    Constraints enforced:
      - Operator pool: a task is deferred until enough operators are free.
      - SMED teams: concurrent setups are capped at `smed_teams`.
      - Machine availability: machines cannot overlap tasks.
      - Setup time: incurred whenever the tool changes between consecutive jobs
        on the same machine (waived for a machine's very first job).
      - Alternative machines: evaluated in priority order (primary → alt1 → alt2 → alt3).
    """

    def __init__(
        self,
        machines: set,
        tools,
        operators_available: int,
        smed_teams: int,
        setup_time: int,
    ):
        # Machine state: next free minute, active tool, accumulated setup, first-use flag
        self.machines = {
            m: {
                "next_free_time": 0,
                "current_tool": None,
                "setup_time_accumulated": 0,
                "first_use": True,
            }
            for m in machines
        }
        self.tools = {t: False for t in tools}

        self.total_operators = operators_available
        self.operators_in_use: list[float] = []   # completion times of active operators

        self.setup_time = setup_time
        self.smed_teams = smed_teams
        self.setups_in_progress: list[float] = [] # finish times of active setups

        # KPI accumulators
        self.total_lateness = 0
        self.total_setup_time = 0
        self.total_earliness = 0
        self.early_task_count = 0
        self.total_tasks = 0

        # Machine utilisation counters
        self.primary_machine_usage = 0
        self.alternative1_usage = 0
        self.alternative2_usage = 0
        self.alternative3_usage = 0

        # Gantt data: one record per scheduled task
        self.gantt_data: list[dict] = []

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _wait_for_operators(self, required: int, current_time: float) -> float:
        """
        Advance `current_time` until `required` operators become free.
        Operators are released at their task's completion time (earliest-first).
        """
        while len(self.operators_in_use) + required > self.total_operators:
            earliest_release = min(self.operators_in_use)
            self.operators_in_use.remove(earliest_release)
            current_time = max(current_time, earliest_release)
        return current_time

    def _wait_for_smed_team(self, current_time: float) -> float:
        """
        Advance `current_time` until a SMED team slot is available.
        """
        while len(self.setups_in_progress) >= self.smed_teams:
            earliest_finish = min(self.setups_in_progress)
            self.setups_in_progress.remove(earliest_finish)
            current_time = max(current_time, earliest_finish)
        return current_time

    def _compute_setup(self, machine: str, tool: str) -> int:
        """
        Return the setup time for assigning `tool` to `machine`.
        No setup is charged on a machine's first ever use, nor when the tool
        is unchanged from the previous job.
        """
        m = self.machines[machine]
        if m["first_use"] or m["current_tool"] == tool:
            return 0
        return self.setup_time

    # ------------------------------------------------------------------
    # Core scheduling logic
    # ------------------------------------------------------------------

    def assign_task(self, task: dict, current_time: float) -> None:
        """
        Assign a single work order to the best available machine.

        Selection criteria (in priority order):
          1. Earliest finish time
          2. Lowest setup time (tie-break)
          3. Lowest slack — i.e. most urgent order (secondary tie-break)

        The machine priority list is [primary, alt1, alt2, alt3]; a lower-priority
        machine is only chosen if it genuinely yields an earlier completion.
        """
        required_ops = task["Nb_Operator"]

        # Block until operator capacity is available
        current_time = self._wait_for_operators(required_ops, current_time)

        # Block until a SMED team slot is available (needed if a setup will occur)
        current_time = self._wait_for_smed_team(current_time)

        # Evaluate all candidate machines
        best_machine = None
        best_finish = float("inf")
        best_setup = float("inf")
        best_slack = float("inf")
        best_index = 0

        candidates = [
            task["Machine"],
            task["MachAlter1"],
            task["MachAlter2"],
            task["MachAlter3"],
        ]

        for idx, machine in enumerate(candidates):
            if pd.isna(machine) or machine not in self.machines:
                continue

            setup = self._compute_setup(machine, task["Tool"])
            start  = max(current_time, self.machines[machine]["next_free_time"]) + setup
            finish = start + task["Proc_time(min)"]
            slack  = task["Due_date_min"] - finish

            # Update best candidate if this machine is strictly better
            if (
                best_machine is None
                or finish < best_finish
                or (finish == best_finish and setup < best_setup)
                or (finish == best_finish and setup == best_setup and slack < best_slack)
            ):
                best_machine = machine
                best_finish  = finish
                best_setup   = setup
                best_slack   = slack
                best_index   = idx

        if best_machine is None:
            print(f"⚠  No machine available for order {task['OF']} — skipping.")
            return

        # --- Commit the assignment ---
        self.operators_in_use.append(best_finish)

        m_state = self.machines[best_machine]
        m_state["first_use"]              = False
        m_state["next_free_time"]         = best_finish
        m_state["setup_time_accumulated"] += best_setup
        m_state["current_tool"]           = task["Tool"]

        self.total_setup_time += best_setup

        if best_setup > 0:
            # Register the setup's finish time so SMED capacity is correctly tracked
            self.setups_in_progress.append(current_time + best_setup)

        # Update machine utilisation counters
        usage_counters = [
            "primary_machine_usage",
            "alternative1_usage",
            "alternative2_usage",
            "alternative3_usage",
        ]
        setattr(self, usage_counters[best_index],
                getattr(self, usage_counters[best_index]) + 1)

        # KPI updates
        self.total_lateness  += max(0, best_finish - task["Due_date_min"])
        self.total_earliness += max(0, task["Due_date_min"] - best_finish)
        if task["Due_date_min"] > best_finish:
            self.early_task_count += 1
        self.total_tasks += 1

        # Record Gantt entry (start here is after setup, i.e. processing start)
        processing_start = best_finish - task["Proc_time(min)"]
        self.gantt_data.append({
            "OF":      task["OF"],
            "Machine": best_machine,
            "Start":   processing_start,
            "Finish":  best_finish,
            "Setup":   best_setup,
        })

    def execute_schedule(self, task_list: list[dict], scenario_name: str) -> None:
        """
        Run the full scheduling pass for `task_list` under the current parameters.

        Ordering: EDD first, then group by tool within the same due date to
        minimise tool-change setups.
        """
        task_list.sort(key=lambda t: (t["Due_date_min"], t["Tool"]))

        for task in task_list:
            self.assign_task(task, current_time=0)

        # Derived KPIs
        avg_earliness_days = (
            (self.total_earliness / MINUTES_PER_DAY) / self.early_task_count
            if self.early_task_count > 0 else 0
        )

        print(f"\n{'=' * 55}")
        print(f"  Results — {scenario_name}")
        print(f"{'=' * 55}")
        print(f"  Total lateness          : {self.total_lateness:>10,} min")
        print(f"  Total setup time        : {self.total_setup_time:>10,} min")
        print(f"  Avg earliness           : {avg_earliness_days:>10.2f} days")
        print(f"  Work orders processed   : {self.total_tasks:>10}")
        print(f"  Primary machine usage   : {self.primary_machine_usage:>10}")
        print(f"  Alternative 1 usage     : {self.alternative1_usage:>10}")
        print(f"  Alternative 2 usage     : {self.alternative2_usage:>10}")
        print(f"  Alternative 3 usage     : {self.alternative3_usage:>10}")

        self.plot_gantt_chart(scenario_name)

    def plot_gantt_chart(self, scenario_name: str) -> None:
        """
        Render a Gantt chart for the current scenario.
        Processing bars are coloured per task; setup blocks are shown in grey.
        A red dashed line marks the 7-day planning horizon (10 080 min).
        """
        if not self.gantt_data:
            print(f"⚠  No Gantt data available for {scenario_name}.")
            return

        fig, ax = plt.subplots(figsize=(14, 7))

        machines_sorted = sorted({d["Machine"] for d in self.gantt_data})
        machine_index   = {m: i for i, m in enumerate(machines_sorted)}

        for task in self.gantt_data:
            y = machine_index[task["Machine"]]

            # Processing bar
            ax.barh(y=y, width=task["Finish"] - task["Start"],
                    left=task["Start"], height=0.4)

            # Setup bar (grey, semi-transparent)
            if task["Setup"] > 0:
                ax.barh(y=y, width=task["Setup"],
                        left=task["Start"] - task["Setup"],
                        height=0.4, color="grey", alpha=0.45)

        ax.set_yticks(range(len(machines_sorted)))
        ax.set_yticklabels(machines_sorted, fontsize=7)
        ax.set_xlabel("Time (minutes)")
        ax.set_ylabel("Machine")
        ax.set_title(f"Gantt Chart — Production Schedule — {scenario_name}")
        ax.axvline(x=10_080, color="red", linestyle="--",
                   linewidth=1.8, label="7-day horizon")
        ax.legend(loc="upper right")

        plt.tight_layout()
        plt.show()


# ---------------------------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------------------------

def main() -> None:
    # --- Load and prepare data ---
    print("Loading production data …")
    data_final, n_grouped, n_day8 = load_and_prepare_data(INPUT_FILE_PATH)
    print(f"  Work orders after deduplication : {len(data_final)}")
    print(f"  Duplicate rows merged           : {n_grouped}")
    print(f"  Day-8 orders absorbed           : {n_day8}")

    # Collect all machines (primary + alternatives) and tools from the dataset
    machines: set[str] = set(data_final["Machine"].dropna().unique())
    for col in ("MachAlter1", "MachAlter2", "MachAlter3"):
        machines.update(data_final[col].dropna().unique())
    tools = data_final["Tool"].dropna().unique()

    # --- Run each scenario ---
    for scenario_name, params in SCENARIOS.items():
        print(f"\n🚀 Running {scenario_name} …")
        print(
            f"   Operators: {params['n_operators']} | "
            f"SMED teams: {params['n_smed_teams']} | "
            f"Setup time: {params['base_setup_time']} min"
        )

        scheduler = ProductionScheduler(
            machines          = machines,
            tools             = tools,
            operators_available = params["n_operators"],
            smed_teams        = params["n_smed_teams"],
            setup_time        = params["base_setup_time"],
        )

        scheduler.execute_schedule(data_final.to_dict("records"), scenario_name)

        # Build a lookup from work-order ID → Gantt record for fast merging
        gantt_lookup = {entry["OF"]: entry for entry in scheduler.gantt_data}

        # --- Write results back to the corresponding Excel sheet ---
        try:
            df_scenario = pd.read_excel(OUTPUT_FILE_PATH, sheet_name=scenario_name)
        except Exception as exc:
            print(f"  ⚠  Could not read sheet '{scenario_name}' from output file: {exc}")
            continue

        for idx, row in df_scenario.iterrows():
            of_id = row["OF"]
            if of_id not in gantt_lookup:
                continue

            t = gantt_lookup[of_id]
            df_scenario.at[idx, "Machine.1"]        = t["Machine"]
            df_scenario.at[idx, "Start setup"]      = t["Start"] - t["Setup"]
            df_scenario.at[idx, "End setup"]        = t["Start"]
            df_scenario.at[idx, "Start Processing"] = t["Start"]
            df_scenario.at[idx, "End Processing"]   = t["Finish"]
            df_scenario.at[idx, "Tardiness"]        = max(0, t["Finish"] - row["Due_date"] * MINUTES_PER_DAY)
            df_scenario.at[idx, "Setup"]            = t["Setup"]
            df_scenario.at[idx, "Earliness"]        = max(0, row["Due_date"] * MINUTES_PER_DAY - t["Finish"])

        with pd.ExcelWriter(
            OUTPUT_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_scenario.to_excel(writer, sheet_name=scenario_name, index=False)

        print(f"  ✅ Results written to sheet '{scenario_name}'.")


if __name__ == "__main__":
    main()
