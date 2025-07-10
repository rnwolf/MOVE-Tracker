"""This script helps you predict when a project will be finished based on the team's past performance.
It creates a visual "burnup chart" to show the most likely completion date and a range of other possible completion dates,
which helps in understanding the uncertainty in the forecast.

"""

import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from scipy import stats


def calculate_burnup_intersection(completion_data, total_scope, start_date):
    """
    Calculate the intersection between a linear trend line and scope line in a burnup chart.

    This function takes project data (completion dates and the number of items completed)
    and the total number of items to be completed (the "scope"). It then:

     -  Calculates a "trend line" using linear regression to determine the
        project's velocity (the rate at which work is completed).
     -  Projects this trend line into the future to estimate the date
        when all work will be completed.
     -  Returns the projected completion date, the number of days remaining,
        and statistical details about the trend line (slope, R-squared).

    Args:
        completion_data: List of tuples [(date, cumulative_completed), ...]
                        where date is datetime object or string 'YYYY-MM-DD'
        total_scope: Total number of work items in scope
        start_date: Project start date (datetime object or string 'YYYY-MM-DD')

    Returns:
        dict: {
            'projected_completion_date': datetime,
            'days_to_completion': float,
            'slope': float,
            'intercept': float,
            'r_squared': float
        }
    """
    # Convert dates to datetime if they're strings
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, "%Y-%m-%d")

    # Prepare data
    dates = []
    cumulative_completed = []

    for date, completed in completion_data:
        if isinstance(date, str):
            date = datetime.strptime(date, "%Y-%m-%d")
        dates.append(date)
        cumulative_completed.append(completed)

    # Convert dates to days from start
    days_from_start = [(date - start_date).days for date in dates]

    # Perform linear regression
    X = np.array(days_from_start).reshape(-1, 1)
    y = np.array(cumulative_completed)

    model = LinearRegression()
    model.fit(X, y)

    slope = model.coef_[0]
    intercept = model.intercept_
    r_squared = model.score(X, y)

    # Calculate intersection with scope line
    # Solve: slope * x + intercept = total_scope
    # x = (total_scope - intercept) / slope

    if slope <= 0:
        raise ValueError("Negative or zero slope detected. Cannot project completion.")

    days_to_completion = (total_scope - intercept) / slope
    projected_completion_date = start_date + timedelta(days=days_to_completion)

    return {
        "projected_completion_date": projected_completion_date,
        "days_to_completion": days_to_completion,
        "slope": slope,
        "intercept": intercept,
        "r_squared": r_squared,
    }


def plot_burnup_chart(
    completion_data, total_scope, start_date, result, confidence_interval=0.95
):
    """
    Plot the burnup chart with trend line and projection with uncertainty bands.

     It creates a chart that shows:

    - The actual progress of completed work over time.
    - A line representing the total scope of the project.
    - The calculated trend line based on the team's velocity.
    - A projection of when the project is likely to finish.
    - A "confidence interval" which is a shaded area showing the earliest and
        latest likely completion dates, providing a more realistic forecast that accounts for variability

    Args:
        confidence_interval: Confidence level for uncertainty bands (default 0.95 for 95%)
    """
    # Convert dates
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, "%Y-%m-%d")

    dates = []
    cumulative_completed = []

    for date, completed in completion_data:
        if isinstance(date, str):
            date = datetime.strptime(date, "%Y-%m-%d")
        dates.append(date)
        cumulative_completed.append(completed)

    # Calculate uncertainty metrics
    days_from_start = [(date - start_date).days for date in dates]
    X = np.array(days_from_start).reshape(-1, 1)
    y = np.array(cumulative_completed)

    # Calculate residuals and standard error
    y_pred = result["slope"] * np.array(days_from_start) + result["intercept"]
    residuals = y - y_pred
    mse = np.mean(residuals**2)
    std_error = np.sqrt(mse)

    # Calculate confidence interval multiplier (t-distribution)
    from scipy import stats

    n = len(dates)
    degrees_freedom = n - 2
    alpha = 1 - confidence_interval
    t_value = stats.t.ppf(1 - alpha / 2, degrees_freedom)

    # Create the plot
    plt.figure(figsize=(12, 8))

    # Plot actual data points
    plt.plot(
        dates, cumulative_completed, "bo-", label="Actual Completion", markersize=6
    )

    # Plot scope line
    plt.axhline(
        y=total_scope, color="red", linestyle="--", label=f"Total Scope ({total_scope})"
    )

    # Plot trend line
    trend_x = np.array(days_from_start)
    trend_y = result["slope"] * trend_x + result["intercept"]
    plt.plot(
        dates,
        trend_y,
        "g--",
        label=f'Trend Line (slope={result["slope"]:.2f})',
        alpha=0.7,
    )

    # Calculate projection with uncertainty - extend far enough to capture confidence bounds
    projection_date = result["projected_completion_date"]
    last_date = dates[-1]

    # Estimate how far we need to project to capture the upper confidence bound
    # Start with a conservative estimate and extend if needed
    base_projection_days = (projection_date - last_date).days

    # Calculate initial projection to determine confidence band width
    initial_days = max(
        base_projection_days * 2, 30
    )  # At least 30 days or 2x the base projection
    temp_timeline = [last_date + timedelta(days=i) for i in range(initial_days + 1)]
    temp_days_from_start = [(date - start_date).days for date in temp_timeline]

    # Calculate uncertainty for initial projection
    temp_uncertainty_multiplier = np.sqrt(
        1
        + ((np.array(temp_days_from_start) - np.mean(days_from_start)) ** 2)
        / np.sum((np.array(days_from_start) - np.mean(days_from_start)) ** 2)
    )
    temp_margin_of_error = t_value * std_error * temp_uncertainty_multiplier
    temp_projection_y = (
        result["slope"] * np.array(temp_days_from_start) + result["intercept"]
    )
    temp_upper_bound = temp_projection_y + temp_margin_of_error

    # Find where the upper bound would intersect the scope line
    upper_intersection_day = None
    for i in range(len(temp_upper_bound) - 1):
        if (temp_upper_bound[i] <= total_scope <= temp_upper_bound[i + 1]) or (
            temp_upper_bound[i] >= total_scope >= temp_upper_bound[i + 1]
        ):
            # Linear interpolation to find exact intersection
            x1, y1 = temp_days_from_start[i], temp_upper_bound[i]
            x2, y2 = temp_days_from_start[i + 1], temp_upper_bound[i + 1]
            if y2 != y1:  # Avoid division by zero
                upper_intersection_day = x1 + (total_scope - y1) * (x2 - x1) / (y2 - y1)
                break

    # Extend projection to at least the upper bound intersection + some buffer
    if upper_intersection_day:
        projection_days = max(
            int(upper_intersection_day - (last_date - start_date).days) + 10,
            base_projection_days,
        )
    else:
        projection_days = base_projection_days * 3  # Fallback if no intersection found

    # Create extended timeline for projection
    projection_timeline = [
        last_date + timedelta(days=i) for i in range(projection_days + 1)
    ]
    projection_days_from_start = [
        (date - start_date).days for date in projection_timeline
    ]

    # Calculate central projection line
    projection_y = (
        result["slope"] * np.array(projection_days_from_start) + result["intercept"]
    )

    # Calculate uncertainty bands
    # Standard error increases with distance from the data
    uncertainty_multiplier = np.sqrt(
        1
        + ((np.array(projection_days_from_start) - np.mean(days_from_start)) ** 2)
        / np.sum((np.array(days_from_start) - np.mean(days_from_start)) ** 2)
    )

    margin_of_error = t_value * std_error * uncertainty_multiplier
    upper_bound = projection_y + margin_of_error
    lower_bound = projection_y - margin_of_error

    # Plot projection with uncertainty bands
    plt.plot(
        projection_timeline,
        projection_y,
        "r:",
        linewidth=2,
        label="Projection to Completion",
    )
    plt.fill_between(
        projection_timeline,
        lower_bound,
        upper_bound,
        alpha=0.3,
        color="red",
        label=f"{int(confidence_interval*100)}% Confidence Interval",
    )

    # Calculate uncertainty range for completion date
    # Find where upper and lower bounds intersect with scope line
    scope_intersections = []
    for bound_y in [lower_bound, upper_bound]:
        # Find intersection with scope line
        for i in range(len(bound_y) - 1):
            if (bound_y[i] <= total_scope <= bound_y[i + 1]) or (
                bound_y[i] >= total_scope >= bound_y[i + 1]
            ):
                # Linear interpolation to find exact intersection
                x1, y1 = projection_days_from_start[i], bound_y[i]
                x2, y2 = projection_days_from_start[i + 1], bound_y[i + 1]
                if y2 != y1:  # Avoid division by zero
                    x_intersect = x1 + (total_scope - y1) * (x2 - x1) / (y2 - y1)
                    scope_intersections.append(start_date + timedelta(days=x_intersect))
                break

    # Mark intersection point and uncertainty range
    plt.plot(
        projection_date, total_scope, "ro", markersize=10, label="Projected Completion"
    )

    # Add uncertainty markers on scope line
    if len(scope_intersections) >= 2:
        early_date, late_date = sorted(scope_intersections)
        plt.plot(
            [early_date, late_date],
            [total_scope, total_scope],
            "r-",
            linewidth=4,
            alpha=0.5,
            label="Completion Range",
        )
        plt.plot(
            [early_date, late_date],
            [total_scope, total_scope],
            "r|",
            markersize=15,
            markeredgewidth=2,
        )

    # Formatting
    plt.xlabel("Date")
    plt.ylabel("Cumulative Items Completed")
    plt.title("Burnup Chart with Linear Projection and Uncertainty")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Add text box with results including uncertainty
    early_str = ""
    late_str = ""
    if len(scope_intersections) >= 2:
        early_date, late_date = sorted(scope_intersections)
        early_str = f"Earliest Completion: {early_date.strftime('%Y-%m-%d')}\n"
        late_str = f"Latest Completion: {late_date.strftime('%Y-%m-%d')}\n"
        range_days = (late_date - early_date).days

    textstr = f"""Projected Completion: {projection_date.strftime('%Y-%m-%d')}
{early_str}{late_str}Days to Completion: {result['days_to_completion']:.1f}
Velocity: {result['slope']:.2f} items/day
R² Score: {result['r_squared']:.3f}
Confidence Level: {int(confidence_interval*100)}%"""

    props = dict(boxstyle="round", facecolor="wheat", alpha=0.8)
    plt.text(
        0.02,
        0.98,
        textstr,
        transform=plt.gca().transAxes,
        fontsize=10,
        verticalalignment="top",
        bbox=props,
    )

    plt.show()


# Example usage
if __name__ == "__main__":
    # Sample data: (date, cumulative_completed)
    sample_data = [
        ("2025-01-01", 0),
        ("2025-01-05", 1),
        # ("2025-01-10", 8),
        # ("2025-01-15", 12),
        # ("2025-01-20", 18),
        # ("2025-01-25", 22),
        # ("2025-01-30", 28),
        # ("2025-02-04", 33),
        # ("2025-02-09", 38),
        # ("2025-02-14", 42),
    ]

    total_scope = 4
    start_date = "2025-01-01"

    # Calculate intersection
    result = calculate_burnup_intersection(sample_data, total_scope, start_date)

    # Print results
    print("Burnup Chart Analysis Results:")
    print(
        f"Projected Completion Date: {result['projected_completion_date'].strftime('%Y-%m-%d')}"
    )
    print(f"Days to Completion: {result['days_to_completion']:.1f}")
    print(f"Velocity (slope): {result['slope']:.2f} items/day")
    print(f"Y-intercept: {result['intercept']:.2f}")
    print(f"R² Score: {result['r_squared']:.3f}")

    # Plot the chart
    plot_burnup_chart(sample_data, total_scope, start_date, result)
