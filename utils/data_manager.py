import streamlit as st
import pandas as pd
from typing import Dict, Any
from utils.collaboration_manager import CollaborationManager

class DataManager:
    @staticmethod
    def initialize_session_state():
        """세션 상태 초기화"""
        if 'excel_data' not in st.session_state:
            st.session_state.excel_data = {}
        if 'current_sheet' not in st.session_state:
            st.session_state.current_sheet = None
        if 'file_uploaded' not in st.session_state:
            st.session_state.file_uploaded = False
        if 'filename' not in st.session_state:
            st.session_state.filename = None
        if 'project_id' not in st.session_state:
            st.session_state.project_id = None
        if 'user_id' not in st.session_state:
            st.session_state.user_id = None
        if 'collaboration_manager' not in st.session_state:
            st.session_state.collaboration_manager = CollaborationManager()
        if 'current_version' not in st.session_state:
            st.session_state.current_version = 0
        if 'is_collaborative' not in st.session_state:
            st.session_state.is_collaborative = False
    
    @staticmethod
    def save_excel_data(excel_data: Dict[str, pd.DataFrame], filename: str):
        """엑셀 데이터를 세션 상태에 저장"""
        st.session_state.excel_data = excel_data
        st.session_state.filename = filename
        st.session_state.file_uploaded = True
        
        # 첫 번째 시트를 기본으로 설정
        if excel_data:
            st.session_state.current_sheet = list(excel_data.keys())[0]
    
    @staticmethod
    def create_collaborative_project(excel_data: Dict[str, pd.DataFrame], filename: str):
        """공동 편집 프로젝트 생성"""
        collaboration_manager = st.session_state.collaboration_manager
        project_id = collaboration_manager.create_project(excel_data, filename)
        
        st.session_state.project_id = project_id
        st.session_state.is_collaborative = True
        st.session_state.user_id = collaboration_manager._generate_user_id()
        
        # 프로젝트에 참여
        collaboration_manager.join_project(project_id, st.session_state.user_id)
        
        return project_id
    
    @staticmethod
    def join_collaborative_project(project_id: str):
        """공동 편집 프로젝트에 참여"""
        collaboration_manager = st.session_state.collaboration_manager
        
        if collaboration_manager.join_project(project_id):
            st.session_state.user_id = collaboration_manager._generate_user_id()
            project_data = collaboration_manager.get_project_data(project_id)
            
            if project_data:
                st.session_state.project_id = project_id
                st.session_state.is_collaborative = True
                st.session_state.excel_data = project_data["excel_data"]
                st.session_state.filename = project_data["metadata"]["filename"]
                st.session_state.file_uploaded = True
                st.session_state.current_version = project_data["metadata"]["version"]
                
                # 첫 번째 시트를 기본으로 설정
                if project_data["excel_data"]:
                    st.session_state.current_sheet = list(project_data["excel_data"].keys())[0]
                
                return True
        return False
    
    @staticmethod
    def sync_collaborative_data():
        """공동 편집 데이터 동기화"""
        if not st.session_state.is_collaborative or not st.session_state.project_id:
            return False
        
        collaboration_manager = st.session_state.collaboration_manager
        project_data = collaboration_manager.get_project_data(st.session_state.project_id)
        
        if project_data:
            # 버전 체크
            server_version = project_data["metadata"]["version"]
            if server_version > st.session_state.current_version:
                # 새 버전이 있으면 업데이트
                st.session_state.excel_data = project_data["excel_data"]
                st.session_state.current_version = server_version
                return True
        return False
    
    @staticmethod
    def update_collaborative_data():
        """공동 편집 데이터 업데이트"""
        if not st.session_state.is_collaborative or not st.session_state.project_id:
            return False
        
        collaboration_manager = st.session_state.collaboration_manager
        success = collaboration_manager.update_project_data(
            st.session_state.project_id,
            st.session_state.excel_data,
            st.session_state.user_id
        )
        
        if success:
            # 버전 업데이트
            st.session_state.current_version += 1
        
        return success
    
    @staticmethod
    def get_active_users():
        """활성 사용자 목록 가져오기"""
        if not st.session_state.is_collaborative or not st.session_state.project_id:
            return []
        
        collaboration_manager = st.session_state.collaboration_manager
        return collaboration_manager.get_active_users(st.session_state.project_id)
    
    @staticmethod
    def update_sheet_data(sheet_name: str, updated_df: pd.DataFrame):
        """특정 시트의 데이터 업데이트"""
        if 'excel_data' in st.session_state:
            st.session_state.excel_data[sheet_name] = updated_df
            
            # 공동 편집 모드에서는 자동으로 서버에 업데이트
            if st.session_state.is_collaborative:
                DataManager.update_collaborative_data()
    
    @staticmethod
    def get_current_data():
        """현재 선택된 시트의 데이터 반환"""
        if st.session_state.current_sheet and st.session_state.excel_data:
            return st.session_state.excel_data.get(st.session_state.current_sheet)
        return None
    
    @staticmethod
    def get_all_data():
        """모든 시트 데이터 반환"""
        return st.session_state.excel_data
    
    @staticmethod
    def clear_data():
        """모든 데이터 초기화"""
        st.session_state.excel_data = {}
        st.session_state.current_sheet = None
        st.session_state.file_uploaded = False
        st.session_state.filename = None
        st.session_state.project_id = None
        st.session_state.user_id = None
        st.session_state.current_version = 0
        st.session_state.is_collaborative = False
