import cv2
import mediapipe as mp
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import tkinter as tk
from threading import Thread
import math
import time

class HandTrackingApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Contrôle du Volume avec la Main")
        self.master.configure(bg="lightblue")

        self.instructions_label = tk.Label(master, text="Utilisez l'index et le pouce pour contrôler le volume.", bg="lightblue")
        self.instructions_label.pack(pady=10)

        self.start_button = tk.Button(master, text="Démarrer", command=self.start_tracking, bg="green", fg="white")
        self.start_button.pack(pady=10)

        self.stop_button = tk.Button(master, text="Arrêter", command=self.stop_tracking, bg="red", fg="white")
        self.stop_button.pack(pady=10)

        self.volume_label = tk.Label(master, text="Volume: 50%", bg="lightblue")
        self.volume_label.pack(pady=10)

        self.canvas = tk.Canvas(master, width=300, height=300, bg="lightgray")
        self.canvas.pack(pady=10)

        self.is_tracking = False
        self.sensitivity = 0.05  # Sensibilité ajustable

    def draw_volume_gauge(self, volume):
        """Dessine un indicateur de volume circulaire."""
        self.canvas.delete("all")
        center_x = 150
        center_y = 150
        radius = 100

        # Dessiner le cercle de fond
        self.canvas.create_oval(center_x - radius, center_y - radius, 
                                center_x + radius, center_y + radius, 
                                outline="black", width=5)

        angle = (volume * 180) - 90
        end_x = center_x + radius * math.cos(math.radians(angle))
        end_y = center_y + radius * math.sin(math.radians(angle))

        self.canvas.create_line(center_x, center_y, end_x, end_y, 
                                fill="blue", width=5)

        self.canvas.create_text(center_x, center_y + 30, 
                                text=f"{int(volume * 100)}%", 
                                font=("Arial", 14, "bold"), fill="black")

    def start_tracking(self):
        self.is_tracking = True
        tracking_thread = Thread(target=self.track_hands)
        tracking_thread.start()

    def stop_tracking(self):
        self.is_tracking = False

    def track_hands(self):
        mp_hands = mp.solutions.hands
        hands = mp_hands.Hands()
        mp_drawing = mp.solutions.drawing_utils

        # Initialisation COM
        CoInitialize()

        # Configuration du volume
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)

        current_volume = volume.GetMasterVolumeLevelScalar()
        last_update_time = time.time()  # Pour contrôler la fréquence de mise à jour

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Erreur lors de l'ouverture de la caméra.")
            return

        while self.is_tracking:
            ret, frame = cap.read()
            if not ret:
                print("Erreur de capture vidéo")
                break

            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = hands.process(frame_rgb)

            if results.multi_hand_landmarks:
                for hand_landmarks in results.multi_hand_landmarks:
                    mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

                    # Récupérer les coordonnées de l'index et du pouce
                    index_finger = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
                    thumb_finger = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]

                    index_x, index_y = index_finger.x, index_finger.y
                    thumb_x, thumb_y = thumb_finger.x, thumb_finger.y

                    # Calculer la distance entre l'index et le pouce
                    distance = math.sqrt((index_x - thumb_x) ** 2 + (index_y - thumb_y) ** 2)

                    # Ajuster le volume selon la distance
                    if distance < self.sensitivity:  # Si les doigts sont proches
                        if time.time() - last_update_time > 0.2:  # Mise à jour toutes les 0.2 secondes
                            current_volume = min(current_volume + 0.01, 1.0)
                            volume.SetMasterVolumeLevelScalar(current_volume, None)
                            last_update_time = time.time()
                    elif distance > self.sensitivity * 2:  # Si les doigts sont éloignés
                        if time.time() - last_update_time > 0.2:
                            current_volume = max(current_volume - 0.01, 0.0)
                            volume.SetMasterVolumeLevelScalar(current_volume, None)
                            last_update_time = time.time()

                    self.volume_label.config(text="Volume: {:.0f}%".format(current_volume * 100))
                    self.draw_volume_gauge(current_volume)

            cv2.imshow('Hand Tracking', frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

        # Libération des ressources COM
        CoUninitialize()

if __name__ == "__main__":
    root = tk.Tk()
    app = HandTrackingApp(root)
    root.mainloop()