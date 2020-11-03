import os

#settings.pyからそのままコピー
BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = '1149=f(e#03_)t@f_@=uoa&lcf#fu%5+_hj*wq!3hyupagj3ec'

#settings.pyからそのままコピー
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

DEBUG = True #ローカルでDebugできるようになります