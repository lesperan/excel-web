import json
import os
import time
import hashlib
from datetime import datetime
from filelock import FileLock
import pandas as pd
from typing import Dict, List, Optional, Tuple
import uuid

class CollaborationManager:
    def __init__(self, project_dir="shared_projects"):
        self.project_dir = project_dir
        self.ensure_project_dir()
    
    def ensure_project_dir(self):
        """프로젝트 디렉토리가 존재하는지 확인하고 생성"""
        if not os.path.exists(self.project_dir):
            os.makedirs(self.project_dir)
    
    def create_project(self, excel_data: Dict[str, pd.DataFrame], filename: str) -> str:
        """새 프로젝트 생성"""
        project_id = self._generate_project_id()
        project_path = os.path.join(self.project_dir, project_id)
        
        # 프로젝트 디렉토리 생성
        os.makedirs(project_path, exist_ok=True)
        
        # 프로젝트 메타데이터 생성
        metadata = {
            "project_id": project_id,
            "filename": filename,
            "created_at": datetime.now().isoformat(),
            "last_modified": datetime.now().isoformat(),
            "active_users": {},
            "version": 1
        }
        
        # 메타데이터 저장
        with open(os.path.join(project_path, "metadata.json"), "w") as f:
            json.dump(metadata, f, indent=2)
        
        # 엑셀 데이터 저장
        self._save_excel_data(project_path, excel_data)
        
        return project_id
    
    def join_project(self, project_id: str, user_id: str = None) -> bool:
        """프로젝트에 참여"""
        if not user_id:
            user_id = self._generate_user_id()
        
        project_path = os.path.join(self.project_dir, project_id)
        if not os.path.exists(project_path):
            return False
        
        # 사용자 활동 기록
        self._update_user_activity(project_id, user_id)
        return True
    
    def get_project_data(self, project_id: str) -> Optional[Dict]:
        """프로젝트 데이터 가져오기"""
        project_path = os.path.join(self.project_dir, project_id)
        if not os.path.exists(project_path):
            return None
        
        try:
            # 메타데이터 로드
            with open(os.path.join(project_path, "metadata.json"), "r") as f:
                metadata = json.load(f)
            
            # 엑셀 데이터 로드
            excel_data = self._load_excel_data(project_path)
            
            return {
                "metadata": metadata,
                "excel_data": excel_data
            }
        except Exception as e:
            print(f"프로젝트 데이터 로드 오류: {e}")
            return None
    
    def update_project_data(self, project_id: str, excel_data: Dict[str, pd.DataFrame], user_id: str = None) -> bool:
        """프로젝트 데이터 업데이트"""
        project_path = os.path.join(self.project_dir, project_id)
        if not os.path.exists(project_path):
            return False
        
        lock_file = os.path.join(project_path, "update.lock")
        
        try:
            with FileLock(lock_file):
                # 메타데이터 업데이트
                metadata_path = os.path.join(project_path, "metadata.json")
                with open(metadata_path, "r") as f:
                    metadata = json.load(f)
                
                metadata["last_modified"] = datetime.now().isoformat()
                metadata["version"] += 1
                
                if user_id:
                    metadata["active_users"][user_id] = datetime.now().isoformat()
                
                with open(metadata_path, "w") as f:
                    json.dump(metadata, f, indent=2)
                
                # 엑셀 데이터 저장
                self._save_excel_data(project_path, excel_data)
                
                return True
        except Exception as e:
            print(f"프로젝트 데이터 업데이트 오류: {e}")
            return False
    
    def get_active_users(self, project_id: str) -> List[Dict]:
        """활성 사용자 목록 가져오기"""
        project_data = self.get_project_data(project_id)
        if not project_data:
            return []
        
        active_users = []
        current_time = datetime.now()
        
        for user_id, last_activity in project_data["metadata"]["active_users"].items():
            last_activity_time = datetime.fromisoformat(last_activity)
            time_diff = (current_time - last_activity_time).total_seconds()
            
            # 5분 이내 활동한 사용자만 활성으로 간주
            if time_diff < 300:
                active_users.append({
                    "user_id": user_id,
                    "last_activity": last_activity
                })
        
        return active_users
    
    def list_projects(self) -> List[Dict]:
        """모든 프로젝트 목록 가져오기"""
        projects = []
        
        if not os.path.exists(self.project_dir):
            return projects
        
        for project_id in os.listdir(self.project_dir):
            project_path = os.path.join(self.project_dir, project_id)
            if os.path.isdir(project_path):
                metadata_path = os.path.join(project_path, "metadata.json")
                if os.path.exists(metadata_path):
                    try:
                        with open(metadata_path, "r") as f:
                            metadata = json.load(f)
                        projects.append({
                            "project_id": project_id,
                            "filename": metadata.get("filename", "Unknown"),
                            "created_at": metadata.get("created_at", ""),
                            "last_modified": metadata.get("last_modified", ""),
                            "active_users_count": len(self.get_active_users(project_id))
                        })
                    except:
                        continue
        
        return sorted(projects, key=lambda x: x["last_modified"], reverse=True)
    
    def _generate_project_id(self) -> str:
        """프로젝트 ID 생성"""
        return str(uuid.uuid4())[:8]
    
    def _generate_user_id(self) -> str:
        """사용자 ID 생성"""
        return f"user_{str(uuid.uuid4())[:8]}"
    
    def _save_excel_data(self, project_path: str, excel_data: Dict[str, pd.DataFrame]):
        """엑셀 데이터를 JSON 형태로 저장"""
        data_to_save = {}
        for sheet_name, df in excel_data.items():
            data_to_save[sheet_name] = df.to_dict('records')
        
        with open(os.path.join(project_path, "data.json"), "w") as f:
            json.dump(data_to_save, f, indent=2, ensure_ascii=False)
    
    def _load_excel_data(self, project_path: str) -> Dict[str, pd.DataFrame]:
        """저장된 엑셀 데이터 로드"""
        data_path = os.path.join(project_path, "data.json")
        if not os.path.exists(data_path):
            return {}
        
        with open(data_path, "r") as f:
            data = json.load(f)
        
        excel_data = {}
        for sheet_name, records in data.items():
            excel_data[sheet_name] = pd.DataFrame(records)
        
        return excel_data
    
    def _update_user_activity(self, project_id: str, user_id: str):
        """사용자 활동 시간 업데이트"""
        project_path = os.path.join(self.project_dir, project_id)
        metadata_path = os.path.join(project_path, "metadata.json")
        
        try:
            with open(metadata_path, "r") as f:
                metadata = json.load(f)
            
            metadata["active_users"][user_id] = datetime.now().isoformat()
            
            with open(metadata_path, "w") as f:
                json.dump(metadata, f, indent=2)
        except Exception as e:
            print(f"사용자 활동 업데이트 오류: {e}")
    
    def get_project_version(self, project_id: str) -> int:
        """프로젝트 버전 가져오기"""
        project_data = self.get_project_data(project_id)
        if project_data:
            return project_data["metadata"].get("version", 1)
        return 0
