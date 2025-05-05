import pandas as pd
import numpy as np
from scipy.signal import hilbert
import matplotlib.pyplot as plt
import io
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas


file_path = "2. MTBE Motors List and Data.xlsx"
motor_names = [ "EA0117AM", "EA0119AM", "EA0121AM", "EA0222AM", "EA0225AM","EA0901AM","EA0901BM","EC1101AM","EC1101BM","EC1101CM",
              "EC1101DM", "KB0701AM", "KC0101AM","PC0101CM","PC0105BM","PC0222BM", "PC0601BM","PC1007BM","PC1106M","PC1108BM"]

#  Hilbert transform function
def hilbert_analysis(signal):
    analytic = hilbert(signal)
    envelope = np.abs(analytic)
    phase = np.unwrap(np.angle(analytic))
    freq = np.gradient(phase) / (2 * np.pi)
    return envelope, freq

# Health Check
def is_healthy(freq, threshold=0.0001):
    return np.std(freq) < threshold

# Processing and generation of graphs for each engine
def process_motor(df, motor_name):
    df['A_avg'] = (df['A(A) Max'] + df['A(A) Min']) / 2
    df['B_avg'] = (df['B(A) Max'] + df['B(A) Min']) / 2
    df['C_avg'] = (df['C(A) Max'] + df['C(A) Min']) / 2

    envelope_A, freq_A = hilbert_analysis(df['A_avg'])
    envelope_B, freq_B = hilbert_analysis(df['B_avg'])
    envelope_C, freq_C = hilbert_analysis(df['C_avg'])

    std_A = np.std(freq_A)
    std_B = np.std(freq_B)
    std_C = np.std(freq_C)

    health_A = is_healthy(freq_A)
    health_B = is_healthy(freq_B)
    health_C = is_healthy(freq_C)

    global_health = "Healthy" if all([health_A, health_B, health_C]) else "Unhealthy"

    health_results = {
        'Phase': ['A', 'B', 'C', 'Global'],
        'Santé': ['Healthy' if health_A else 'Unhealthy',
                  'Healthy' if health_B else 'Unhealthy',
                  'Healthy' if health_C else 'Unhealthy',
                  global_health],
        'Standard deviation frequency': [std_A, std_B, std_C, np.nan]
    }
    health_df = pd.DataFrame(health_results)

    fig, axs = plt.subplots(2, 1, figsize=(14, 10))
    axs[0].plot(envelope_A, label='Envelope A')
    axs[0].plot(envelope_B, label='Envelope B')
    axs[0].plot(envelope_C, label='Envelope C')
    axs[0].set_title(f" Signal envelope - Phases A, B, C - {motor_name}")
    axs[0].legend()
    axs[0].grid(True)

    axs[1].plot(freq_A, label='Frequency A')
    axs[1].plot(freq_B, label='Frequency B')
    axs[1].plot(freq_C, label='Frequency C')
    axs[1].set_title(f"instant frequency - Phases A, B, C - {motor_name}")
    axs[1].legend()
    axs[1].grid(True)

    plt.tight_layout()

    img_stream = io.BytesIO()
    canvas = FigureCanvas(fig)
    canvas.print_png(img_stream)

    return health_df, img_stream

#  Traitement avec gestion des erreurs
output_file_path = "MTBE_Motors_with_health_results.xlsx"
xls = pd.ExcelFile(file_path)
available_sheets = xls.sheet_names

print(" Sheets available in the file :")
print(available_sheets)

all_health_results = []

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    workbook = writer.book

    for motor_name in motor_names:
        print(f"\n Treatment of: {motor_name}")
        if motor_name not in available_sheets:
            print(f" Leaf '{motor_name}' not found. Engine ignored.")
            continue

        try:
            df_motor = pd.read_excel(file_path, sheet_name=motor_name)
            health_df, img_stream = process_motor(df_motor, motor_name)

            health_df.to_excel(writer, sheet_name=motor_name, index=False)

            worksheet = workbook.add_worksheet(f'Graphiques_{motor_name}')
            worksheet.insert_image('A1', 'graph.png', {'image_data': img_stream})

            # Accumulation des résultats globaux
            temp_df = health_df.copy()
            temp_df.insert(0, 'Motor', "")  # Insère une colonne 'Moteur' en première position
            temp_df.loc[temp_df.index[0], 'Motor'] = motor_name  # Écrit le nom du moteur uniquement sur la première ligne
            all_health_results.append(temp_df)


        except Exception as e:
            print(f" Error when processing {motor_name} : {e}")
            continue

    # Résumé global
    if all_health_results:
        final_health_df = pd.concat(all_health_results, ignore_index=True)
        final_health_df.to_excel(writer, sheet_name=' Global results', index=False)
    else:
        print("No valid engines processed, no global results.")

print("\n Processing complete. File generated :", output_file_path)
