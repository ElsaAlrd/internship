import pandas as pd
import numpy as np
from scipy.signal import hilbert
import matplotlib.pyplot as plt
import io
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

# source file
file_path = "1. PPTSB Motors List and Data.xlsx"

# Motors list
motor_names = [
    "E22802-01", "E22802-02", "E22802-03", "E22802-04", "E22802-05",
    "E22802-06", "E22802-08", "E22807-01", "E22807-02", "E22812-01",
    "E22812-02", "E22902-01", "E22902-02", "E22902-04", "E22905-01",
    "E22905-02", "PM22803A", "PM22808B", "PM22902B"
]

# hilbery analysis
def hilbert_analysis(signal):
    analytic = hilbert(signal)              # z(t) = x(t) + j·Hilbert(x(t))
    envelope = np.abs(analytic)             # |z(t)| = √(x² + Hilbert(x)²)
    phase = np.unwrap(np.angle(analytic))   # instant phase
    freq = np.gradient(phase) / (2 * np.pi) # instant freq
    return envelope, freq

# Health check
def is_healthy(freq, threshold=0.0001):
    return np.std(freq) < threshold         # standard deviation< threshold

# Motor processing
def process_motor(df, motor_name):
  # calculation of average signals
    df['A_avg'] = (df['A(A) Max'] + df['A(A) Min']) / 2
    df['B_avg'] = (df['B(A) Max'] + df['B(A) Min']) / 2
    df['C_avg'] = (df['C(A) Max'] + df['C(A) Min']) / 2


# hilbery analysis
    envelope_A, freq_A = hilbert_analysis(df['A_avg'])
    envelope_B, freq_B = hilbert_analysis(df['B_avg'])
    envelope_C, freq_C = hilbert_analysis(df['C_avg'])

# health check
    std_A = np.std(freq_A)
    std_B = np.std(freq_B)
    std_C = np.std(freq_C)

    health_A = is_healthy(freq_A)
    health_B = is_healthy(freq_B)
    health_C = is_healthy(freq_C)

    global_health = "Healthy" if all([health_A, health_B, health_C]) else "Unhealthy"

    # Ne mettre le nom du moteur que sur la première ligne (Phase A)
    moteur_column = [motor_name, "", "", ""]

    # result health
    health_df = pd.DataFrame({
        'Moteur': moteur_column,
        'Phase': ['A', 'B', 'C', 'Global'],
        'Santé': ['Healthy' if health_A else 'Unhealthy',
                  'Healthy' if health_B else 'Unhealthy',
                  'Healthy' if health_C else 'Unhealthy',
                  global_health],
        'Standard deviation frequency': [std_A, std_B, std_C, np.nan]
    })

    # Genrate graph
    fig, axs = plt.subplots(2, 1, figsize=(14, 10))
    # envelopes
    axs[0].plot(envelope_A, label='Envelope A')
    axs[0].plot(envelope_B, label='Envelope B')
    axs[0].plot(envelope_C, label='Envelope C')
    axs[0].set_title(f"Signal envelope - Phases A, B, C - {motor_name}")
    axs[0].legend()
    axs[0].grid(True)

# Frequency
    axs[1].plot(freq_A, label='Frequency A')
    axs[1].plot(freq_B, label='Frequency B')
    axs[1].plot(freq_C, label=' Frequency')
    axs[1].set_title(f" Instant frequency  - Phases A, B, C - {motor_name}")
    axs[1].legend()
    axs[1].grid(True)

    plt.tight_layout()
    img_stream = io.BytesIO()
    canvas = FigureCanvas(fig)
    canvas.print_png(img_stream)
    plt.close(fig)

    return health_df, img_stream

# Excel file generation
output_file_path = "PPTSB_Motors_with_health_results.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    all_health_results = []

    for motor_name in motor_names:
        df_motor = pd.read_excel(file_path, sheet_name=motor_name)
        health_df, img_stream = process_motor(df_motor, motor_name)

        # Feuille spécifique moteur
        health_df.to_excel(writer, sheet_name=motor_name, index=False)

        # Graph
        workbook = writer.book
        worksheet = workbook.add_worksheet(f'Graphiques_{motor_name}')
        worksheet.insert_image('A1', 'graph.png', {'image_data': img_stream})

        all_health_results.append(health_df)

    # Nettoyage et fusion des résultats globaux
    cleaned_results = [df.dropna(how='all') for df in all_health_results]
    final_health_df = pd.concat(cleaned_results, ignore_index=True)
    final_health_df.to_excel(writer, sheet_name='Global results', index=False)

print("\n Processing complete. Results and graphs saved in PPTSB_Motors_with_health_results.xlsx")
