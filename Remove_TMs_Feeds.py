import pandas as pd
import re
import os

# ====================================================================
# CONFIGURATION
# ====================================================================
# Danh sách từ khóa bị cấm (blacklist)
# Việc kiểm tra không phân biệt chữ hoa chữ thường.
BLACKLIST_KEYWORDS = [
    "Whyitsme", "Cottagecore", "Trump", "Biden", "Reggae", "Smoke Daddy", "Celtic Cross", "Bob Marley", "Family Guy", "Gay Cat", "Gay Trash", "Fishy", "Venom", "Boba", "BSN", "Uterus", "Van Gogh",
    "CARHARTT", "Nonni", "Kangaroo", "Tuxedo", "Dibble", "Dabble", "Oh ship", "COHIBA", "Jurassic", "Jeep", "Jeeps", "Adventure Before Dementia", "antisocial", "anti social", "Cobra", "Python",
    "Spirit Halloween", "Got Titties", "Le Tits Now", "Mack Trucks", "V-buck", "V buck", "Vbuck", "World Traveler", "Rollerblade", "Black Lives Matter", "Just The Tip", "In My Defense", "Sleep Token",
    "U.S.Army", "US Army", "Crazy Chicken Lady", "Christmas In July", "Grill Sergeant", "Ducks Unlimited", "SOTALLY Tober", "Birds aren't Real", "Pickleballer", "Quaker", "Vampire Mansion",
    "Lampoon's", "Lampoons", "Lampoon", "krampus", "griswold", "Brainrot", "Disney", "Marvel", "Star Wars", "Music Television", "MTV", "Fender", "Nightmare Before Christmas", "Life is Good",
    "WWE", "NFL", "NBA", "Robux", "ASPCA", "Alpha Wolf", "Milkshake", "milk_shake", "Costume Agent", "La Colombe", "Tesla", "LeBron", "Seuss", "Grinch", "Peanuts", "Pixar", "InGENIUS",
]

# Từ điển các từ khóa cần thay thế
# Key là từ cũ, Value là từ mới.
REPLACEMENT_KEYWORDS = {
    "Guess": "Funny",
    "Rubiks": "Cube",
    "Jockey": "Funny",
    "comica": "Funny",
    "Sakura": "Flower",
    "Superhero": "Heroes",
    "Yeti": "Bigfoot",
    "Beast": "Strong",
    "Diesel": "Handyman",
    "K-Pop": "Korean Music",
    "Kpop": "Korean Music",
    "Frisbee": "Sport",
    "Coach": "Fun",
    "KOOZIE": "Drinking",
    "Prosecco": "Drinking",
    "Craftsman": "Handyman",
    "Pajama": "Costume",
    "Pajamas": "Costume",
    "Shark Week": "Shark Lovers",
    "BANNED": "Reading Lover",
    "Arcade": "Game Machine",
    "Ducky": "Duck Lovers",
    "Skittles": "Fruit Candy",
    "Akita": "Dog",
    "Lucky Charms": "Lucky Gifts",
    "Little Trees": "Trees Lover",
    "Fallout": "Radiation",
    "Fuck": "Fck",
    "Halls": "Holidays",
    "Mr. Christmas": "Couple Christmas",
    "Mr Christmas": "Couple Christmas",
    "Busch": "Funny",
}

# Tên các cột cần kiểm tra trong file Excel
# Sử dụng tên tag (hàng thứ ba của header) để đảm bảo tính duy nhất.
COLUMNS_TO_CHECK = [
    "item_name",
    "product_description",
    "bullet_point1",
    "bullet_point2",
    "bullet_point3",
    "bullet_point4",
    "bullet_point5",
    "generic_keywords"
]

# ====================================================================
# FUNCTIONS
# ====================================================================

def contains_blacklist_keyword_with_info(text, blacklist):
    """
    Kiểm tra xem một chuỗi có chứa bất kỳ từ khóa nào trong blacklist không
    và trả về từ khóa đầu tiên tìm thấy.
    """
    if pd.isna(text):
        return None
    
    text_lower = str(text).lower()
    for keyword in blacklist:
        if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', text_lower):
            return keyword
    return None

def replace_keywords(text, replacements):
    """
    Thay thế các từ khóa trong một chuỗi bằng các từ khóa tương ứng.
    """
    if pd.isna(text):
        return ""
    
    modified_text = str(text)
    for old_word, new_word in replacements.items():
        modified_text = re.sub(r'\b' + re.escape(old_word) + r'\b', new_word, modified_text, flags=re.IGNORECASE)
    return modified_text

def process_excel_file(input_filename):
    """
    Hàm chính để xử lý file Excel.
    """
    if not input_filename.endswith('.xlsx'):
        input_filename += '.xlsx'
    
    output_filename = os.path.splitext(input_filename)[0] + "_processed.xlsx"

    try:
        print(f"Đang đọc file: {input_filename}...")
        full_df = pd.read_excel(input_filename, header=None)
        
        header_rows = full_df.iloc[:3]
        
        df_data = full_df.iloc[3:].copy()
        df_data.columns = full_df.iloc[2].tolist()
        df_data.index = range(len(df_data))
        
        initial_rows = len(df_data)
        print(f"Tổng số hàng dữ liệu ban đầu: {initial_rows}")

        # Bước 1: Xóa các hàng chứa từ khóa blacklist và ghi log
        rows_to_delete = []
        deleted_log = []

        for index, row in df_data.iterrows():
            row_deleted = False
            for col in COLUMNS_TO_CHECK:
                if col in row:
                    found_keyword = contains_blacklist_keyword_with_info(row[col], BLACKLIST_KEYWORDS)
                    if found_keyword:
                        # Ghi log chi tiết
                        original_row_number = index + 4
                        deleted_log.append({
                            "Hàng": original_row_number,
                            "Từ khóa bị cấm": found_keyword,
                            "Cột": col
                        })
                        rows_to_delete.append(index)
                        row_deleted = True
                        break # Chuyển sang hàng tiếp theo nếu đã tìm thấy từ khóa
            
        df_cleaned = df_data.drop(rows_to_delete)
        
        deleted_count = len(rows_to_delete)
        print(f"----------------------------------------------------")
        print(f"TÓM TẮT XỬ LÝ")
        print(f"Tổng số hàng dữ liệu ban đầu: {initial_rows}")
        print(f"Tổng số hàng bị xóa: {deleted_count}")
        
        if deleted_count > 0:
            print(f"\nCHI TIẾT CÁC HÀNG BỊ XÓA:")
            for log_entry in deleted_log:
                print(f"- Hàng {log_entry['Hàng']}: Bị xóa vì từ khóa '{log_entry['Từ khóa bị cấm']}' trong cột '{log_entry['Cột']}'")
        else:
            print(f"Không có hàng nào bị xóa do chứa từ khóa blacklist.")
        
        print(f"----------------------------------------------------")

        # Bước 2: Thay thế các từ khóa trên dữ liệu đã được làm sạch
        for column in COLUMNS_TO_CHECK:
            if column in df_cleaned.columns:
                print(f"Đang thay thế từ khóa trong cột '{column}'...")
                df_cleaned[column] = df_cleaned[column].apply(
                    lambda text: replace_keywords(text, REPLACEMENT_KEYWORDS)
                )
        
        # Tạo một writer object để ghi dữ liệu
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # Ghi 3 hàng header gốc vào file mới
            header_rows.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
            
            # Ghi dữ liệu đã được xử lý vào sau 3 hàng header
            df_cleaned.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=3)

        final_rows = len(df_cleaned)
        print(f"\nTổng số hàng dữ liệu sau khi xử lý: {final_rows}")
        print(f"Đã lưu kết quả vào file: {output_filename}")

    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file '{input_filename}'. Vui lòng kiểm tra lại tên file và đường dẫn.")
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")

# ====================================================================
# MAIN EXECUTION
# ====================================================================
if __name__ == "__main__":
    file_to_process = input("Vui lòng nhập tên file Excel cần xử lý (ví dụ: my_file): ")
    process_excel_file(file_to_process)