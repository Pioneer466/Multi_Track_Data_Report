import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import glob

sns.set_theme(style="whitegrid", context="talk")
plt.rcParams["figure.dpi"] = 120
plt.rcParams["axes.titlesize"] = 18
plt.rcParams["axes.labelsize"] = 14
plt.rcParams["xtick.labelsize"] = 12
plt.rcParams["ytick.labelsize"] = 12

class Colors:
    HEADER = "\033[95m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"
    END = "\033[0m"


def auto_load_excel(folder="data"):
    files = glob.glob(folder + "/student_grades_*.xlsx")

    if not files:
        raise FileNotFoundError("No matching Excel files found in /data")
    
    files.sort()
    latest = files[-1]
    return pd.read_excel(latest, sheet_name=None)


def load_and_clean_excel():

    # Loading all sheets
    sheets = auto_load_excel()

    all_dfs = []

    for track_name, df in sheets.items():
        df["Track"] = track_name

        # Clean special values
        df = df.replace(["Waived", "N/A", "NA", ""], pd.NA)

        # Normalize everything into numeric formats
        numeric_cols = ["Math", "English", "Science", "History", "Attendance (%)", "ProjectScore"]
        for col in numeric_cols:
            df[col] = df[col].astype(str).str.extract(r'(\d+\.?\d*)', expand=False)
            df[col] = pd.to_numeric(df[col], errors="coerce")

        # Turn IncomeStudent and Passed into numeric formats
        df["IncomeStudent"] = df["IncomeStudent"].astype(bool)

        df["Passed (Y/N)"] = df["Passed (Y/N)"].astype(str).str.strip().str.upper()
        df["Passed (Y/N)"] = df["Passed (Y/N)"].map({"Y": True, "N": False})

        all_dfs.append(df)

    final_df = pd.concat(all_dfs, ignore_index=True)

    return final_df


def compute_statistics(df):

    stats = {}

    # Track-level statistics
    track_stats = df.groupby("Track").agg(
        Students=("Track", "count"),
        MathAvg=("Math", "mean"),
        EnglishAvg=("English", "mean"),
        ScienceAvg=("Science", "mean"),
        HistoryAvg=("History", "mean"),
        AttendanceAvg=("Attendance (%)", "mean"),
        ProjectAvg=("ProjectScore", "mean"),
        PassRate=("Passed (Y/N)", "mean")
    )

    stats["track"] = track_stats


    # Cohort-level statistics
    cohort_stats = df.groupby("Cohort").agg(
        Students=("Track", "count"),
        MathAvg=("Math", "mean"),
        EnglishAvg=("English", "mean"),
        ScienceAvg=("Science", "mean"),
        HistoryAvg=("History", "mean"),
        AttendanceAvg=("Attendance (%)", "mean"),
        ProjectAvg=("ProjectScore", "mean"),
        PassRate=("Passed (Y/N)", "mean")
        )
    stats["cohort"] = cohort_stats


    # Income-student comparison
    income_stats = df.groupby("IncomeStudent").agg(
        Students=("Track", "count"),
        MathAvg=("Math", "mean"),
        EnglishAvg=("English", "mean"),
        ScienceAvg=("Science", "mean"),
        HistoryAvg=("History", "mean"),
        AttendanceAvg=("Attendance (%)", "mean"),
        ProjectAvg=("ProjectScore", "mean"),
        PassRate=("Passed (Y/N)", "mean")
        )
    stats["income"] = income_stats

    

    # Inter-track comparison

    # History grades distribution
    stats["history_by_track"] = df.groupby("Track")["History"].apply(list)

    # Comparing Math scores
    math_comparison = df.groupby("Track")["Math"].mean()
    stats["math_comparison"] = math_comparison

    # Correlation between Attendance and ProjectScore
    corr_list = {}
    for track in df["Track"].unique():
        subdf = df[df["Track"] == track][["Attendance (%)", "ProjectScore"]]
        corr = subdf.corr().iloc[0, 1]    # corr(A, P)
        corr_list[track] = corr

    stats["attendance_project_corr"] = pd.Series(corr_list)

     # Global stats with all tracks combined
    global_stats = df.agg({
        "Math": "mean",
        "English": "mean",
        "Science": "mean",
        "History": "mean",
        "Attendance (%)": "mean",
        "ProjectScore": "mean",
        "Passed (Y/N)": "mean"
    })

    stats["global"] = global_stats

    return stats


def generate_visuals(df, output_folder="output"):

    # Create output folder if not exists
    os.makedirs(output_folder, exist_ok=True)

    PALETTE = sns.color_palette("viridis")
    tracks = df["Track"].unique()

    # Histogram of History by track

    for i, track in enumerate(tracks):
        plt.figure(figsize=(9, 5))
        subset = df[df["Track"] == track]["History"]

        sns.histplot(
            subset,
            kde=True,
            bins=15,
            color=PALETTE[i % len(PALETTE)]
        )
        plt.title(f"History Grade Distribution – {track}")
        plt.xlabel("History Score")
        plt.ylabel("Frequency")
        plt.grid(alpha=0.3)

        plt.tight_layout()
        plt.savefig(f"{output_folder}/history_distribution_{track}.png")
        plt.close()


    # Box plot of history by grade

    plt.figure(figsize=(10, 6))
    sns.boxplot(
        data=df,
        x="Track",
        y="History",
        color=PALETTE[2]
    )
    plt.title("History Score Distribution by Track")
    plt.xlabel("Track")
    plt.ylabel("History Score")
    plt.grid(alpha=0.3)

    plt.tight_layout()
    plt.savefig(f"{output_folder}/history_boxplot.png")
    plt.close()


    # Barplot of avg math by track

    math_means = df.groupby("Track")["Math"].mean()

    plt.figure(figsize=(10, 5))
    sns.barplot(
        x=math_means.index,
        y=math_means.values,
        color=PALETTE[4]   
    )
    plt.title("Average Math Score by Track")
    plt.ylabel("Average Math Score")
    plt.xlabel("Track")
    plt.grid(axis='y', alpha=0.3)

    plt.tight_layout()
    plt.savefig(f"{output_folder}/math_comparison.png")
    plt.close()


    # Correlation attendance vs Project score per track

    for i, track in enumerate(tracks):
        plt.figure(figsize=(8, 5))
        subset = df[df["Track"] == track]

        sns.scatterplot(
            data=subset,
            x="Attendance (%)",
            y="ProjectScore",
            color=PALETTE[i % len(PALETTE)],
            s=60
        )

        sns.regplot(
            data=subset,
            x="Attendance (%)",
            y="ProjectScore",
            scatter=False,
            color=PALETTE[(i + 2) % len(PALETTE)],
            line_kws={"linewidth": 2}
        )

        plt.title(f"Attendance vs Project Score – {track}")
        plt.xlabel("Attendance (%)")
        plt.ylabel("Project Score")
        plt.grid(alpha=0.3)

        plt.tight_layout()
        plt.savefig(f"{output_folder}/attendance_project_corr_{track}.png")
        plt.close()


def export_outputs(df, stats, output_folder="output"):

    os.makedirs(output_folder, exist_ok=True)

    # Export cleaned merged dataset
    cleaned_path = os.path.join(output_folder, "cleaned_dataset.csv")
    df.to_csv(cleaned_path, index=False)
    print(f"Cleaned dataset exported to: {cleaned_path}")


    # Export all statistics in an excel file
    stats_path = os.path.join(output_folder, "summary_statistics.xlsx")

    with pd.ExcelWriter(stats_path, engine="openpyxl") as writer:
        for key, table in stats.items():

            # Convert Series to DataFrame for consistency
            if isinstance(table, pd.Series):
                table = table.to_frame()
            table.to_excel(writer, sheet_name=key)

    print(f"Statistics exported to: {stats_path}")

    print("\nAll outputs exported successfully!")


def generate_performance_alerts(stats, mode="track"):
    print()
    print(Colors.BOLD + Colors.RED + "=== PERFORMANCE ALERTS ===" + Colors.END)
    print()

    if mode == "track":
        table = stats["track"]

        for track, row in table.iterrows():
            math_avg = row["MathAvg"]
            pass_rate = row["PassRate"]

            if math_avg < 70:
                print(f"{Colors.RED}⚠️ ALERT: Track '{track}' has LOW Math performance (avg = {math_avg:.1f}){Colors.END}")

            if pass_rate < 0.6:
                print(f"{Colors.RED}⚠️ ALERT: Track '{track}' has LOW Pass Rate ({pass_rate*100:.1f} %){Colors.END}")

        print("\nDone.")

    elif mode == "cohort":
        table = stats["cohort"]

        for cohort, row in table.iterrows():
            math_avg = row["MathAvg"]
            pass_rate = row["PassRate"]

            if math_avg < 70:
                print(f"{Colors.RED}⚠️ ALERT: Cohort '{cohort}' has LOW Math performance (avg = {math_avg:.1f}){Colors.END}")

            if pass_rate < 0.6:
                print(f"{Colors.RED}⚠️ ALERT: Cohort '{cohort}' has LOW Pass Rate ({pass_rate*100:.1f} %){Colors.END}")

        print("\nDone.")

    else:
        print(Colors.YELLOW + "Invalid mode." + Colors.END)


#----------#
### CLI ###
#----------#

def clear():
    os.system("cls" if os.name == "nt" else "clear")


def main_menu():
    while True:
        clear()
        print(Colors.BOLD + Colors.CYAN + "=== STUDENT ANALYTICS MAIN MENU ===" + Colors.END)
        print()
        print(Colors.GREEN + "[1] Track-level analysis" + Colors.END)
        print(Colors.GREEN + "[2] Cohort-level analysis" + Colors.END)
        print(Colors.GREEN + "[3] IncomeStudent analysis" + Colors.END)
        print(Colors.GREEN + "[4] Generate visuals" + Colors.END)
        print(Colors.GREEN + "[5] Export dashboard report" + Colors.END)
        print(Colors.GREEN + "[6] Performance alerts" + Colors.END)
        print(Colors.RED + "[0] Quit" + Colors.END)

        choice = input("\nEnter your choice: ")

        if choice == "1":
            track_submenu()
        elif choice == "2":
            cohort_submenu()
        elif choice == "3":
            income_submenu()
        elif choice == "4":
            visuals_submenu()
        elif choice == "5":
            export_dashboard_submenu()
        elif choice == "6":
            performance_alerts_submenu()
        elif choice == "0":
            clear()
            print("Goodbye!")
            break
        else:
            input(Colors.RED + "Invalid choice. Press ENTER to continue." + Colors.END)


def track_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- TRACK ANALYSIS ---" + Colors.END)
        print()
        print("[1] Show full track statistics")
        print("[2] Compare Math between tracks")
        print("[3] Show Attendance vs Project correlation by track")
        print("[4] Show History distributions")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            print(stats["track"])
            input("\nPress ENTER to continue.")
        elif choice == "2":
            print(stats["math_comparison"])
            input("\nPress ENTER to continue.")
        elif choice == "3":
            print(stats["attendance_project_corr"])
            input("\nPress ENTER to continue.")
        elif choice == "4":
            print(stats["history_by_track"])
            input("\nPress ENTER to continue.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER to continue.")


def cohort_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- COHORT ANALYSIS ---" + Colors.END)
        print()
        print("[1] Show cohort statistics")
        print("[2] Show cohort pass rates only")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            print(stats["cohort"])
            input("\nPress ENTER to continue.")
        elif choice == "2":
            print(stats["cohort"]["PassRate"])
            input("\nPress ENTER to continue.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER to continue.")


def income_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- INCOME STUDENT ANALYSIS ---" + Colors.END)
        print()
        print("[1] Show Income vs Non-Income statistics")
        print("[2] Show pass rates only")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            print(stats["income"])
            input("\nPress ENTER to continue.")
        elif choice == "2":
            print(stats["income"]["PassRate"])
            input("\nPress ENTER to continue.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER to continue.")


def visuals_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- VISUALS MENU ---" + Colors.END)
        print()
        print("[1] Generate all visuals")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            generate_visuals(df)
            input("\nVisuals generated! Press ENTER.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER to continue.")


def export_dashboard_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- EXPORT DASHBOARD ---" + Colors.END)
        print()
        print("[1] Export basic Excel summary")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            export_outputs(df, stats)
            input("\nBasic summary exported. Press ENTER.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER.")


def performance_alerts_submenu():
    while True:
        clear()
        print(Colors.BOLD + Colors.BLUE + "--- PERFORMANCE ALERTS ---" + Colors.END)
        print()
        print("[1] Show track alerts")
        print("[2] Show cohort alerts")
        print("[0] Back to main menu")

        choice = input("\nChoice: ")

        if choice == "1":
            generate_performance_alerts(stats)
            input("\nPress ENTER.")
        elif choice == "2":
            generate_performance_alerts(stats, mode="cohort")
            input("\nPress ENTER.")
        elif choice == "0":
            break
        else:
            input("Invalid choice. Press ENTER.")


def main():
    global df, stats
    df = load_and_clean_excel()
    stats = compute_statistics(df)
    main_menu()


if __name__ == "__main__":
    main()


