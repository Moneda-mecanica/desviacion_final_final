# -*- coding: utf-8 -*-
"""
Created on Thu Oct  3 14:24:06 2024

@author: lportilla
"""



from app_dash import app

if __name__ == "__main__":
    #threading.Timer(1, open_browser).start()
    app.run_server(debug=True)