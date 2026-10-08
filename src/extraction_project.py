"""Persistent project attribution for extracted attendance batches."""
import json
from pathlib import Path


def write_batch_project(folder, project, source_filename):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    data={'project_id':project['project_id'],'project_name':project['display_name'],'source_filename':source_filename}
    (folder/'.batch-project.json').write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')


def read_batch_project(folder):
    try:
        data=json.loads((Path(folder)/'.batch-project.json').read_text(encoding='utf-8'))
        return {k:str(data.get(k,'')) for k in ('project_id','project_name','source_filename')}
    except (OSError,ValueError,TypeError,AttributeError):
        return {'project_id':'','project_name':'','source_filename':''}
