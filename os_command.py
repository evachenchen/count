import sys
from streamlit import cli as stcli
import streamlit as st
import web_page


def main():
    web_page.function()


if __name__ == '__main__':
    if st._is_running_with_streamlit:
        main()
    else:
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())
