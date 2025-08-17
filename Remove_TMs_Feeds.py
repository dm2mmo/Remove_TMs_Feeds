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
    "Spirit Halloween", "Got Titties", "Le Tits Now", "Mack Trucks", "V-buck", "V buck", "Vbuck", "World Traveler", "Rollerblade", "Black Lives Matter", "Just The Tip", "In My Defense", "Van Gogh",
    "U.S.Army", "US Army", "Crazy Chicken Lady", "Christmas In July", "Grill Sergeant", "Ducks Unlimited", "SOTALLY Tober", "Birds aren't Real", "Pickleballer", "Quaker", "Vampire Mansion",
    "Lampoon's", "Lampoons", "Lampoon", "krampus", "griswold", "Brainrot", "Disney", "Marvel", "Star Wars", "Music Television", "MTV", "Fender", "Nightmare Before Christmas", "Life is Good",
    "WWE", "NFL", "NBA", "Robux", "ASPCA", "Alpha Wolf",
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

def contains_blacklist_keyword(text, blacklist):
    """
    Kiểm tra xem một chuỗi có chứa bất kỳ từ khóa nào trong blacklist không.
    """
    if pd.isna(text):
        return False
    
    text_lower = str(text).lower()
    for keyword in blacklist:
        if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', text_lower):
            return True
    return False

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
        # Đọc toàn bộ file Excel mà không chỉ định hàng header
        print(f"Đang đọc file: {input_filename}...")
        full_df = pd.read_excel(input_filename, header=None)
        
        # Tách 3 hàng header đầu tiên
        header_rows = full_df.iloc[:3]
        
        # Tách phần dữ liệu và gán tên cột từ hàng thứ 3 (chỉ mục 2)
        df_data = full_df.iloc[3:].copy()
        df_data.columns = full_df.iloc[2].tolist()
        
        initial_rows = len(df_data)
        print(f"Tổng số hàng dữ liệu ban đầu: {initial_rows}")

        # Bước 1: Xóa các hàng chứa từ khóa blacklist
        rows_to_delete = df_data.apply(
            lambda row: any(contains_blacklist_keyword(row[col], BLACKLIST_KEYWORDS) for col in COLUMNS_TO_CHECK if col in row),
            axis=1
        )
        df_cleaned = df_data[~rows_to_delete].copy()
        
        deleted_rows = initial_rows - len(df_cleaned)
        print(f"Đã xóa {deleted_rows} hàng chứa từ khóa blacklist.")

        # Bước 2: Thay thế các từ khóa
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
            # header=False để không ghi lại tên cột, startrow=3 để bắt đầu từ hàng thứ 4
            df_cleaned.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=3)

        final_rows = len(df_cleaned)
        print(f"Tổng số hàng dữ liệu sau khi xử lý: {final_rows}")
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