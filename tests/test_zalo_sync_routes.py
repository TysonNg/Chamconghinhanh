import pytest
from unittest.mock import patch
from src.app import app


def test_zalo_sync_config_get():
    client = app.test_client()
    mock_data = {
        'success': True,
        'data': {
            'enabled': True,
            'scheduleTime': '22:00',
            'lookbackDays': 10,
            'mappings': []
        }
    }
    with patch('src.app._proxy_zalo', return_value=(mock_data, 200)) as mock_proxy:
        res = client.get('/api/zalo/sync/config')
        assert res.status_code == 200
        data = res.get_json()
        assert data['success'] is True
        assert data['data']['scheduleTime'] == '22:00'
        mock_proxy.assert_called_once_with('/api/sync/config')


def test_zalo_sync_config_post():
    client = app.test_client()
    payload = {'enabled': True, 'scheduleTime': '21:30', 'lookbackDays': 7}
    mock_data = {'success': True, 'data': payload}
    with patch('src.app._proxy_zalo', return_value=(mock_data, 200)) as mock_proxy:
        res = client.post('/api/zalo/sync/config', json=payload)
        assert res.status_code == 200
        data = res.get_json()
        assert data['success'] is True
        mock_proxy.assert_called_once_with('/api/sync/config', method='POST', data=payload)


def test_zalo_sync_gaps():
    client = app.test_client()
    mock_gaps = {
        'success': True,
        'data': {
            'projectName': 'Dự Án Test',
            'missingDates': ['2026-09-28'],
            'hasMissing': True
        }
    }
    with patch('src.app._proxy_zalo', return_value=(mock_gaps, 200)) as mock_proxy:
        res = client.get('/api/zalo/sync/gaps?projectName=D%E1%BB%B1%20%C3%81n%20Test&lookbackDays=10')
        assert res.status_code == 200
        data = res.get_json()
        assert data['success'] is True
        assert data['data']['hasMissing'] is True
        mock_proxy.assert_called_once()
        call_arg = mock_proxy.call_args[0][0]
        assert 'D%C6%B0%CC%A3%20A%CC%81n%20Test' in call_arg or 'D%E1%BB%B1%20%C3%81n%20Test' in call_arg


def test_zalo_sync_run():
    client = app.test_client()
    payload = {'backfillMissing': True}
    mock_res = {'success': True, 'message': 'Đã bắt đầu tiến trình đồng bộ'}
    with patch('src.app._proxy_zalo', return_value=(mock_res, 200)) as mock_proxy:
        res = client.post('/api/zalo/sync/run', json=payload)
        assert res.status_code == 200
        data = res.get_json()
        assert data['success'] is True
        mock_proxy.assert_called_once_with('/api/sync/run', method='POST', data=payload)
