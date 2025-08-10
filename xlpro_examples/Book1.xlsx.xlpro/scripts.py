

import scipy
from scipy.interpolate import CubicSpline
from pathlib import Path
from pandas import pd

def main():
    
    wd = Path(__file__).parent

    dir_data = wd / "../Engineering Data/Raw"

    dfs = []
    fstems = []

    for fp in dir_data.glob("*.csv"):
        df = pd.read_csv(fp)
        df.columns = ["strain", "stress"]
        dfs.append(df)
        fstems.append(fp.stem)

    width = 12.5
    thickness = 3
    area = width * thickness # mm^2
    length = 50 # mm


    def add_multifreq_noise(y, amp=0.01, num_freqs=5, base_freq=50):
        """Adds summed sine wave noise at multiple frequencies """
        n = len(y)
        x = np.linspace(0, 1, n)
        noise = np.zeros(n)
        
        for i in range(1, num_freqs + 1):
            freq = base_freq * i * np.random.uniform(0.8, 1.5)
            phase = np.random.uniform(0, 2 * np.pi)
            noise += np.sin(2 * np.pi * freq * x + phase)
        
        # Normalize to desired amplitude
        noise = amp * noise / np.max(np.abs(noise))
        
        return y + noise


    new_dfs = []

    matplotlib.use("TkAgg")
    for df in dfs:
        displacement = (df["strain"] * length).values
        force = (df["stress"] * area).values
        force = np.hstack([[0.0], force])
        displacement = np.hstack([[0.0], displacement])
        # force[0] = 0.0
        # displacement[0] = 0.0
        spline = CubicSpline(displacement, force)

        xdata = np.linspace(displacement[0], displacement[-1], 1001)
        ydata = spline(xdata)

        ydata_noisy = add_multifreq_noise(np.copy(ydata), amp=10, num_freqs=10, base_freq=500)

        new_df = pd.DataFrame(np.vstack((xdata, ydata_noisy)).T, columns=["Displacement (mm)", "Force (N)"])
        new_dfs.append(new_df)
        # plt.plot(xdata, ydata_noisy)
        # plt.show(block=False)
        pass
    # plt.show()

    dir_out = wd / "../Engineering Data/Modified"
    dir_out.mkdir(exist_ok=True)

    for i, df in enumerate(new_dfs):
        stem = fstems[i]
        df.to_csv(dir_out / f"{stem}.csv")

    

    pass

if __name__ == "__main__":
    main()
