

def code_response(code=1, msg='', data:dict={}):
    from django.http.response import JsonResponse
    ret = {
        'rescode': code,
        'resmsg': msg,
    }
    ret.update(data)
    return JsonResponse(ret)


def generate_jwt_token(openid:str='undefined_wx_openid', encode_to_str=True):
    import jwt # pip install pyjwt
    from tools_common import get_timestamp

    JWT_SECRET = 'dkdll893hj938h42h829h'
    EXPIRE_SECONDS = '7000'
    payload = {
        'uid': 'uid_abc',
        'expire_time': get_timestamp() + EXPIRE_SECONDS,
    }
    token = jwt.encode(payload, JWT_SECRET, algorithm='HS256')  # decoded = jwt.decode(token, secret, algorithms='HS256')
    if encode_to_str:
        token = str(token, encoding='utf-8')
    return token


def parse_jwt_token(token, key=None):
    '''
        if verify passed, return payload
        if key, return payload[key]
        else, return None
    '''
    import jwt # pip install pyjwt
    from tools_common import get_timestamp
    try:
        JWT_SECRET = 'dkdll893hj938h42h829h'
        payload = jwt.decode(token, JWT_SECRET, algorithms='HS256')
        cur_timestamp = get_timestamp()
        if cur_timestamp > payload['expire_time']:
            raise Exception('login expired')
        if key:
            return payload.get(key)
        else:
            return payload
    except Exception as e:
        print(f'verify token failed: {e}')
    return None

