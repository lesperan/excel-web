import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from utils.excel_handler import ExcelHandler
from utils.data_manager import DataManager
import io
import time

# 페이지 설정
st.set_page_config(
    page_title="웹 엑셀 편집기",
    page_icon="📊",
    layout="wide"
)

def show_collaboration_sidebar():
    """공동 편집 관련 사이드바"""
    st.sidebar.header("👥 공동 편집")
    
    # 현재 프로젝트 상태
    if st.session_state.is_collaborative:
        st.sidebar.success(f"🔗 프로젝트 ID: {st.session_state.project_id}")
        st.sidebar.info(f"👤 사용자 ID: {st.session_state.user_id[:8]}...")
        
        # 활성 사용자 표시
        active_users = DataManager.get_active_users()
        if active_users:
            st.sidebar.subheader("🟢 활성 사용자")
            for user in active_users:
                user_display = user["user_id"][:8]
                if user["user_id"] == st.session_state.user_id:
                    user_display += " (나)"
                st.sidebar.text(f"• {user_display}")
        
        # 동기화 버튼
        if st.sidebar.button("🔄 데이터 동기화"):
            if DataManager.sync_collaborative_data():
                st.sidebar.success("✅ 데이터가 동기화되었습니다!")
                st.rerun()
            else:
                st.sidebar.info("ℹ️ 이미 최신 데이터입니다.")
    
    else:
        # 프로젝트 생성
        st.sidebar.subheader("🆕 새 공동 편집 프로젝트")
        if st.session_state.file_uploaded:
            if st.sidebar.button("📤 공동 편집 프로젝트 생성"):
                project_id = DataManager.create_collaborative_project(
                    st.session_state.excel_data,
                    st.session_state.filename
                )
                st.sidebar.success(f"✅ 프로젝트 생성 완료!\nID: {project_id}")
                st.rerun()
        else:
            st.sidebar.info("먼저 엑셀 파일을 업로드하세요.")
        
        # 프로젝트 참여
        st.sidebar.subheader("🚪 프로젝트 참여")
        project_id_input = st.sidebar.text_input("프로젝트 ID 입력")
        if st.sidebar.button("참여하기"):
            if project_id_input:
                if DataManager.join_collaborative_project(project_id_input):
                    st.sidebar.success("✅ 프로젝트에 참여했습니다!")
                    st.rerun()
                else:
                    st.sidebar.error("❌ 프로젝트를 찾을 수 없습니다.")
            else:
                st.sidebar.error("프로젝트 ID를 입력하세요.")
    
    # 프로젝트 목록
    st.sidebar.subheader("📋 프로젝트 목록")
    collaboration_manager = st.session_state.collaboration_manager
    projects = collaboration_manager.list_projects()
    
    if projects:
        for project in projects[:5]:  # 최근 5개만 표시
            with st.sidebar.expander(f"📁 {project['filename'][:20]}..."):
                st.text(f"ID: {project['project_id']}")
                st.text(f"수정: {project['last_modified'][:16]}")
                st.text(f"활성 사용자: {project['active_users_count']}명")
                if st.button(f"참여", key=f"join_{project['project_id']}"):
                    if DataManager.join_collaborative_project(project['project_id']):
                        st.success("참여 완료!")
                        st.rerun()
    else:
        st.sidebar.info("생성된 프로젝트가 없습니다.")

def main():
    st.title("📊 웹 엑셀 편집기 (공동 편집)")
    st.markdown("엑셀 파일을 업로드하고 웹에서 공동으로 편집해보세요!")
    
    # 세션 상태 초기화
    DataManager.initialize_session_state()
    
    # 자동 새로고침을 위한 플레이스홀더
    if st.session_state.is_collaborative:
        # 30초마다 자동 동기화
        if 'last_sync_time' not in st.session_state:
            st.session_state.last_sync_time = time.time()
        
        current_time = time.time()
        if current_time - st.session_state.last_sync_time > 30:
            if DataManager.sync_collaborative_data():
                st.session_state.last_sync_time = current_time
                st.rerun()
            st.session_state.last_sync_time = current_time
    
    # 사이드바 - 파일 업로드 및 관리
    with st.sidebar:
        st.header("📁 파일 관리")
        
        # 파일 업로드
        uploaded_file = st.file_uploader(
            "엑셀 파일을 업로드하세요",
            type=['xlsx', 'xls'],
            help="Excel 파일만 업로드 가능합니다"
        )
        
        if uploaded_file is not None:
            try:
                # 엑셀 파일 읽기
                excel_data = ExcelHandler.read_excel(uploaded_file)
                DataManager.save_excel_data(excel_data, uploaded_file.name)
                st.success(f"✅ '{uploaded_file.name}' 파일이 업로드되었습니다!")
                
            except Exception as e:
                st.error(f"❌ 파일 업로드 오류: {str(e)}")
        
        # 현재 파일 정보
        if st.session_state.file_uploaded:
            if st.session_state.is_collaborative:
                st.info(f"📄 공동 편집 중: {st.session_state.filename}")
            else:
                st.info(f"📄 현재 파일: {st.session_state.filename}")
            
            # 시트 선택
            sheet_names = list(st.session_state.excel_data.keys())
            if len(sheet_names) > 1:
                st.session_state.current_sheet = st.selectbox(
                    "시트 선택",
                    sheet_names,
                    index=sheet_names.index(st.session_state.current_sheet) if st.session_state.current_sheet in sheet_names else 0
                )
            
            # 다운로드 버튼
            if st.button("📥 엑셀 다운로드"):
                try:
                    excel_bytes = ExcelHandler.dataframe_to_excel(st.session_state.excel_data)
                    st.download_button(
                        label="다운로드",
                        data=excel_bytes,
                        file_name=f"edited_{st.session_state.filename}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ 다운로드 오류: {str(e)}")
            
            # 데이터 초기화 버튼
            if st.button("🗑️ 데이터 초기화"):
                DataManager.clear_data()
                st.rerun()
        
        # 구분선
        st.divider()
        
        # 공동 편집 관련 UI
        show_collaboration_sidebar()
    
    # 메인 영역 - 데이터 표시 및 편집
    if st.session_state.file_uploaded:
        current_data = DataManager.get_current_data()
        
        if current_data is not None:
            # 헤더 정보
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader(f"📋 시트: {st.session_state.current_sheet}")
                if st.session_state.is_collaborative:
                    st.caption(f"🔗 공동 편집 모드 | 버전: {st.session_state.current_version}")
            
            with col2:
                if st.session_state.is_collaborative:
                    # 실시간 동기화 상태 표시
                    if st.button("🔄 새로고침"):
                        if DataManager.sync_collaborative_data():
                            st.success("동기화 완료!")
                            st.rerun()
            
            # 데이터 정보
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("행 수", len(current_data))
            with col2:
                st.metric("열 수", len(current_data.columns))
            with col3:
                st.metric("시트 수", len(st.session_state.excel_data))
            with col4:
                if st.session_state.is_collaborative:
                    active_users_count = len(DataManager.get_active_users())
                    st.metric("활성 사용자", active_users_count)
            
            # AgGrid 설정
            gb = GridOptionsBuilder.from_dataframe(current_data)
            gb.configure_default_column(
                editable=True, 
                resizable=True,
                sortable=True,
                filter=True
            )
            gb.configure_grid_options(
                domLayout='normal',
                suppressRowDeselection=True,
                suppressCellSelection=False,
                enableRangeSelection=True
            )
            
            grid_options = gb.build()
            
            # AgGrid 표시
            st.markdown("### 📝 데이터 편집")
            if st.session_state.is_collaborative:
                st.markdown("🤝 **공동 편집 모드**: 변경사항이 자동으로 다른 사용자들과 동기화됩니다.")
            else:
                st.markdown("셀을 클릭하여 직접 편집할 수 있습니다.")
            
            grid_response = AgGrid(
                current_data,
                gridOptions=grid_options,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                update_mode=GridUpdateMode.MODEL_CHANGED,
                fit_columns_on_grid_load=True,
                height=400,
                reload_data=False,
                key=f"grid_{st.session_state.current_sheet}_{st.session_state.current_version}"
            )
            
            # 변경사항 자동 저장
            if grid_response['data'] is not None:
                updated_df = pd.DataFrame(grid_response['data'])
                
                # 데이터가 실제로 변경되었는지 확인
                if not updated_df.equals(current_data):
                    DataManager.update_sheet_data(st.session_state.current_sheet, updated_df)
                    
                    if st.session_state.is_collaborative:
                        st.success("✅ 변경사항이 저장되고 동기화되었습니다!")
                        time.sleep(1)  # 잠시 대기 후 새로고침
                        st.rerun()
            
            # 실시간 활동 표시
            if st.session_state.is_collaborative:
                with st.expander("👥 실시간 공동 편집 상태"):
                    active_users = DataManager.get_active_users()
                    if active_users:
                        st.markdown("**현재 활성 사용자:**")
                        for user in active_users:
                            user_display = user["user_id"][:8]
                            if user["user_id"] == st.session_state.user_id:
                                st.markdown(f"- 🟢 {user_display} (나)")
                            else:
                                st.markdown(f"- 🔵 {user_display}")
                        
                        st.markdown(f"**마지막 업데이트:** {time.strftime('%Y-%m-%d %H:%M:%S')}")
                    else:
                        st.info("현재 활성 사용자가 없습니다.")
            
            # 데이터 미리보기
            with st.expander("🔍 데이터 미리보기 (처음 10행)"):
                st.dataframe(current_data.head(10))
        
        else:
            st.warning("⚠️ 선택된 시트의 데이터가 없습니다.")
    
    else:
        # 파일이 업로드되지 않은 경우
        st.info("👈 사이드바에서 엑셀 파일을 업로드하거나 공동 편집 프로젝트에 참여해주세요.")
        
        # 사용법 안내
        st.markdown("""
        ### 📖 사용법
        
        #### 🔹 개별 편집
        1. **파일 업로드**: 사이드바에서 엑셀 파일(.xlsx, .xls)을 업로드하세요
        2. **시트 선택**: 여러 시트가 있는 경우 원하는 시트를 선택하세요
        3. **데이터 편집**: 표의 셀을 클릭하여 직접 편집할 수 있습니다
        4. **다운로드**: 편집이 완료되면 사이드바에서 다운로드하세요
        
        #### 🔹 공동 편집
        1. **프로젝트 생성**: 파일 업로드 후 "공동 편집 프로젝트 생성" 버튼 클릭
        2. **프로젝트 공유**: 생성된 프로젝트 ID를 다른 사용자들과 공유
        3. **프로젝트 참여**: 프로젝트 ID를 입력하여 공동 편집에 참여
        4. **실시간 편집**: 모든 변경사항이 실시간으로 다른 사용자들과 동기화됩니다
        
        ### ✨ 주요 기능
        - 📁 엑셀 파일 업로드 및 다운로드
        - 📊 웹에서 실시간 데이터 편집
        - 📋 다중 시트 지원
        - 🤝 **실시간 공동 편집**
        - 👥 **활성 사용자 표시**
        - 🔄 **자동 데이터 동기화**
        - 🔍 데이터 미리보기
        """)
        
        # 공동 편집 데모 안내
        st.markdown("""
        ---
        ### 🎯 공동 편집 테스트 방법
        
        1. **브라우저 창을 2개 열어보세요** (또는 시크릿 모드 사용)
        2. **첫 번째 창에서** 엑셀 파일을 업로드하고 공동 편집 프로젝트를 생성하세요
        3. **두 번째 창에서** 생성된 프로젝트 ID로 참여하세요
        4. **양쪽 창에서** 데이터를 편집해보세요 - 실시간으로 동기화됩니다!
        
        💡 **팁**: 30초마다 자동 동기화되며, 수동으로 "새로고침" 버튼을 클릭할 수도 있습니다.
        """)

if __name__ == "__main__":
    main()
