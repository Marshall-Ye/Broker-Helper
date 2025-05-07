class ClientConfig(object):
    PUBLIC_KEY = (
        '{"kty":"RSA",'
        '"n":"/1rhVahzh1mQskp86dXBnaSHQBUnJxBiZpAPkW42bmk",'
        '"e":"AQAB",'
        '"alg":"RS256",'
        '"kid":"/1rhVahzh1mQskp86dXBnaSHQBUnJxBiZpAPkW42bmk"}'
    )
    APP_NAME     = "GABrokerHelper"
    COMPANY_NAME = "Golden Arcus"
    HTTP_TIMEOUT = 30
    MAX_DOWNLOAD_RETRIES = 3
    UPDATE_URLS  = [
        "https://github.com/Marshall-Ye/Broker-Helper/releases/download/"
    ]
