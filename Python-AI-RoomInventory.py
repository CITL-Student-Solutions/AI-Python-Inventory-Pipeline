import pandas as pd

class RoomDeviceInfo:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = self.load_data()

    def load_data(self):
        try:
            df = pd.read_excel(self.file_path)
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
            return df
        except Exception as e:
            print(f"Error loading file: {e}")
            return None

    def search_info(self, room=None, device=None):
        if self.data is None:
            return "Data not loaded properly."

        query_str = ""
        if room:
            query_str += f"(room == '{room}')"
        if device:
            if query_str:
                query_str += " & "
            query_str += f"(device == '{device}')"

        if query_str:
            result = self.data.query(query_str)
            if not result.empty:
                return result
            else:
                return f"No information found for room '{room}' and device '{device}'."
        else:
            return "Please provide a room or device to search."

    def save_filtered_data(self, output_path, room=None, device=None):
        filtered_data = self.search_info(room=room, device=device)
        if isinstance(filtered_data, pd.DataFrame) and not filtered_data.empty:
            try:
                filtered_data.to_excel(output_path, index=False)
                return f"Filtered data saved to {output_path}"
            except Exception as e:
                return f"Error saving file: {e}"
        else:
            return "No data to save."


# Functions to be called by GPT
def search_room_device_info(file_path, room=None, device=None):
    info = RoomDeviceInfo(file_path)
    return info.search_info(room=room, device=device)

def save_room_device_info(file_path, output_path, room=None, device=None):
    info = RoomDeviceInfo(file_path)
    return info.save_filtered_data(output_path, room=room, device=device)
