"""Trang "Tạo phụ đề tiếng Anh (Whisper)": hướng dẫn + nút mở công cụ trên Google Colab (card đồ hoạ miễn phí)."""
import streamlit as st

COLAB_URL = "https://colab.research.google.com/github/kizumaki/editkb/blob/main/Tao_phu_de_Whisper.ipynb"


def render_tab12():
    st.subheader("Tạo phụ đề tiếng Anh bằng Whisper")
    st.markdown("Dùng khi **chưa nhận được script tiếng Anh**: nghe video và tự tạo phụ đề tiếng Anh (.srt). "
                "Chạy trên **Google Colab** — Google cho mượn card đồ hoạ miễn phí nên nhanh hơn nhiều so với chạy trên máy không có card.")
    st.link_button("🚀 Mở công cụ tạo phụ đề (Google Colab)", COLAB_URL, type="primary", use_container_width=True)

    with st.container(border=True):
        st.markdown("#### Cách dùng")
        st.markdown(
            "1. Bấm nút xanh ở trên → đăng nhập **Gmail** nếu được hỏi.\n"
            "2. Bấm menu **Thời gian chạy → Chạy tất cả** (hoặc phím **Ctrl + F9**). Nếu hiện *\"Cảnh báo: sổ tay này không phải do Google tạo\"* → bấm **Vẫn chạy**.\n"
            "3. Chờ 1–2 phút cài đặt, đến khi hiện nút **Chọn tệp** ở cuối trang → chọn **video hoặc file tiếng** (chọn được nhiều file).\n"
            "4. Chờ chạy xong → máy **tự tải về** file `..._EN_whisper.srt` (nhiều file thì 1 file .zip).\n"
            "5. Đưa file .srt vào ScriptPro như bình thường. Khi script chính thức tới, dùng trang **Đối chiếu 2 file tiếng Anh** để xem chỗ khác nhau.")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        st.markdown("#### 🔒 Bảo mật")
        st.markdown("- File tải **thẳng vào phiên làm việc tạm**, KHÔNG lưu lên Google Drive.\n"
                    "- Chỉ **bản sao** gửi lên Colab bị xoá (video gốc trên máy bạn **giữ nguyên**): bản sao video xoá ngay sau khi tách tiếng, bản sao tiếng xoá khi xong; cuối cùng xoá sạch.\n"
                    "- Dự án nào agency **cấm** đưa lên dịch vụ ngoài → **không dùng** công cụ này.")
    with c2, st.container(border=True):
        st.markdown("#### 💡 Mẹo")
        st.markdown("- Gửi **file tiếng** (.mp3, .wav) nhanh hơn nhiều so với video (video nặng nên tải lên lâu).\n"
                    "- Thấy dòng **🐢 \"chạy bằng bộ xử lý thường\"**: Google đang hết card miễn phí → vào **Thời gian chạy → Thay đổi loại thời gian chạy → T4 GPU** rồi chạy lại, hoặc thử lại sau.\n"
                    "- Phụ đề Whisper **không có tên người nói** — phần đó vẫn làm như cũ.")
