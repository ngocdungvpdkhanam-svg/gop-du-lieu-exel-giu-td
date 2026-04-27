import streamlit as st
import pandas as pd
import io

# Cấu hình trang
st.set_page_config(page_title="Excel Merger Pro", layout="wide")

st.title("📂 Công cụ Gộp File Excel")
st.markdown("Tải lên các file Excel có cấu trúc giống nhau để gộp chúng thành một file duy nhất.")

# --- PHẦN 1: CÀI ĐẶT ---
st.sidebar.header("Cài đặt gộp")
keep_header = st.sidebar.checkbox("File có chứa hàng tiêu đề", value=True, 
                                  help="Nếu tích chọn, chương trình sẽ lấy tiêu đề từ file đầu tiên và chỉ lấy dữ liệu từ các file tiếp theo.")

# --- PHẦN 2: TẢI FILE ---
uploaded_files = st.file_uploader(
    "Chọn các file Excel (.xlsx, .xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Đã chọn {len(uploaded_files)} file.")
    
    combined_df = []
    success_count = 0
    
    try:
        for file in uploaded_files:
            # Đọc file
            if keep_header:
                df = pd.read_excel(file)
            else:
                df = pd.read_excel(file, header=None)
            
            combined_df.append(df)
            success_count += 1
        
        # Gộp dữ liệu
        final_df = pd.concat(combined_df, ignore_index=True)
        
        # --- PHẦN 3: HIỂN THỊ XEM TRƯỚC ---
        st.subheader("👀 Xem trước dữ liệu sau khi gộp")
        st.write(f"Tổng số dòng: {final_df.shape[0]} | Tổng số cột: {final_df.shape[1]}")
        st.dataframe(final_df.head(100)) # Hiển thị 100 dòng đầu tiên

        # --- PHẦN 4: XUẤT FILE ---
        # Tạo buffer để lưu file Excel trong bộ nhớ
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, header=keep_header, sheet_name='Sheet1')
            # Không cần gọi writer.save() trong context manager mới
            
        st.divider()
        st.download_button(
            label="📥 Tải file đã gộp về (.xlsx)",
            data=buffer.getvalue(),
            file_name="merged_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Sẵn sàng để tải về!")

    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý: {e}")

else:
    st.warning("Vui lòng tải lên ít nhất một file để bắt đầu.")

# Hướng dẫn nhỏ
with st.expander("Hướng dẫn sử dụng"):
    st.write("""
    1. **Tải file:** Nhấn vào ô 'Browse files' hoặc kéo thả các file Excel vào.
    2. **Tùy chỉnh:** Sử dụng thanh bên trái để chọn có giữ hàng tiêu đề hay không.
    3. **Kiểm tra:** Xem bảng dữ liệu mẫu hiện ra để đảm bảo gộp đúng.
    4. **Tải về:** Nhấn nút 'Tải file đã gộp về' để lưu kết quả.
    """)
