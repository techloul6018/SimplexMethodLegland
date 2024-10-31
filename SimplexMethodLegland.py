
import pandas as pd
import numpy as np

def get_functions_from_xlsx(filename):

    constraints = []
    df = pd.read_excel(filename, header=None)
    rows = df.values.tolist()

    obj_coeffs = [float(x) for x in rows[0] if pd.notna(x)]
    n_vars = len(obj_coeffs)

    for row in rows[1:]:
        coeffs = [float(x) for x in row[:n_vars]]
        inequality = row[n_vars]
        bound = float(row[n_vars + 1])

        if len(coeffs) != n_vars:
            raise ValueError("Mismatch in number of coefficients for constraint.")

        if inequality == "<=":
            constraints.append((coeffs, "<=", bound))
        elif inequality == ">=":
            constraints.append((coeffs, ">=", bound))
        elif inequality == "=":
            constraints.append((coeffs, "=", bound))
        else:
            raise ValueError(f"Unknown inequality: {inequality}")

    return n_vars, obj_coeffs, constraints

def initialize_tab(n_vars, obj_coeffs, constraints):
    num_constraints = len(constraints)
    tableau = np.zeros((num_constraints + 1, n_vars + num_constraints + 1))

    tableau[-1, :n_vars] = -np.array(obj_coeffs)  


    for i, (coeffs, inequality, bound) in enumerate(constraints):
        tableau[i, :n_vars] = coeffs
        if inequality == "<=":
            tableau[i, n_vars + i] = 1  
        elif inequality == ">=":
            tableau[i, n_vars + i] = -1 
        else:
            tableau[i, n_vars + i] = 1  

        tableau[i, -1] = bound

    return tableau

def find_pivot_column(tableau):
    last_row = tableau[-1, :-1]
    pivot_col = np.argmin(last_row)
    return pivot_col if last_row[pivot_col] < 0 else -1

def find_pivot_row(tableau, pivot_col):
    rhs = tableau[:-1, -1]
    pivot_col_values = tableau[:-1, pivot_col]
    ratios = [rhs[i] / pivot_col_values[i] if pivot_col_values[i] > 0 else np.inf for i in range(len(rhs))]
    return np.argmin(ratios) if min(ratios) != np.inf else -1

def pivot(tableau, pivot_row, pivot_col):
    tableau[pivot_row] /= tableau[pivot_row, pivot_col] 
    for i in range(tableau.shape[0]):
        if i != pivot_row:
            tableau[i] -= tableau[i, pivot_col] * tableau[pivot_row]

def print_tab(tableau, pivot_col=None):
    red_color = '\033[91m'
    reset_color = '\033[0m'
    for i, row in enumerate(tableau):
        for j, value in enumerate(row):
            if j == pivot_col:
                print(f"{red_color}{value:>10.2f}{reset_color}", end=" ")
            else:
                print(f"{value:>10.2f}", end=" ")
        print()

def solve_with_simplex(n_vars, obj_coeffs, constraints):
    tableau = initialize_tab(n_vars, obj_coeffs, constraints)
    print("Initial Tableau:")
    print_tab(tableau)
    print()

    while True:
        pivot_col = find_pivot_column(tableau)
        if pivot_col == -1:
            print("Optimal solution found.")
            break

        pivot_row = find_pivot_row(tableau, pivot_col)
        if pivot_row == -1:
            print("The problem is unbounded.")
            return

        pivot(tableau, pivot_row, pivot_col)
        print("Updated Tableau:")
        print_tab(tableau, pivot_col)
        print()

    solution = np.zeros(n_vars)
    for i in range(n_vars):
        column = tableau[:, i]
        if np.sum(column == 1) == 1 and np.sum(column) == 1:
            one_index = np.where(column == 1)[0][0]
            solution[i] = tableau[one_index, -1]

    print(f"Optimal Objective Value: {tableau[-1, -1]}")
    print("Solution:", solution)

def main():
    print("Welcome to the Linear Programming Solver.")
    n_vars, obj_coeffs, constraints = get_functions_from_xlsx(r"C:\OptimizationExercise\SimplexMethod\SimplexFunctions.xlsx")

    print("\nObjective Function:")
    print("Maximize: ", " + ".join(f"{coeff}*x{i+1}" for i, coeff in enumerate(obj_coeffs)))
    print("\nConstraints:")
    for coeffs, inequality, bound in constraints:
        constraint_str = " + ".join(f"{coeff}*x{i+1}" for i, coeff in enumerate(coeffs))
        print(f"{constraint_str} {inequality} {bound}")

    print("\nNote: To modify the objective function or constraints, please edit them in the 'SimplexFunctions.xlsx' file.\n")

    print("\nSolving with Simplex Method by Tableau")
    solve_with_simplex(n_vars, obj_coeffs, constraints)

if __name__ == "__main__":
    main()
