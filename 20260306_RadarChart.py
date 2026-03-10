import os
import numpy as np
import matplotlib.pyplot as plt

data_00 = np.array([40.9903737630347, 16.0020744853234, 17.205808947156, 16.8413992876135, 16.3840393099619])
data_01 = np.array([62.5878878106618, 15.5708875676455, 16.2435821214621, 16.2003036835166, 16.2009393820468])
data_02 = np.array([34.6715084207726, 11.2374208228367, 15.5858438069807, 15.1734909372006, 15.195018530697])
data_03 = np.array([53.3457466904755, 16.0917097803497, 17.5646024844519, 16.8658421333931, 17.1649691635313])
data_norm = np.array([192, 43.3012701892219, 43.3012701892219, 43.672147618493, 43.8861967352935])

values = data_03/data_norm
labels = ["CM Distance", "BSD100", "BSD50", "BSD10", "BSD All"]

# Number of variables
N = len(values)

# Angles for radar chart
angles = np.linspace(0, 2 * np.pi, N, endpoint=False)

# Close the polygon
values_closed = np.concatenate([values, [values[0]]])
angles_closed = np.concatenate([angles, [angles[0]]])

# ---- Calculate enclosed area ----
# For a polygon in polar coordinates:
# A = 1/2 * sum(r_i * r_{i+1} * sin(theta_{i+1} - theta_i))
area = 0.5 * np.sum(
    values * np.roll(values, -1) * np.sin(np.roll(angles, -1) - angles)
)

print(f"Enclosed area = {area:.6f}")

# ---- Plot radar chart ----
fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))

# Start from 12 o'clock
ax.set_theta_offset(np.pi / 2)

# Clockwise direction
ax.set_theta_direction(-1)

# Plot and fill
ax.plot(angles_closed, values_closed, 'o-', linewidth=2)
ax.fill(angles_closed, values_closed, alpha=0.25)

# Labels
ax.set_xticks(angles)
ax.set_xticklabels(labels)

# Make labels blue and bold
for label in ax.get_xticklabels():
    label.set_color("blue")
    label.set_fontweight("bold")

# Move labels outward to avoid overlap
ax.tick_params(axis='x', pad=15)

# Optional radial limit
#ax.set_ylim(0, max(values) * 1.2)
ax.set_ylim(0, 0.45)

plt.title("Radar Plot")
plt.show()