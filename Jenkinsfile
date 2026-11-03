pipeline {
    agent any

    environment {
        APP_DIR = '/opt/scrapper'
        SERVICE = 'scrapper'
        PLAYWRIGHT_BROWSERS_PATH = '/opt/scrapper/browsers'
    }

    options {
        timestamps()
        disableConcurrentBuilds()
        buildDiscarder(logRotator(numToKeepStr: '20'))
    }

    stages {
        stage('Syntax check') {
            steps { sh 'python3 -m py_compile newScrapper.py' }
        }

        stage('Deploy') {
            when { branch 'main' }
            steps {
                sh '''
                    set -e
                    rsync -a newScrapper.py requirements.txt ${APP_DIR}/

                    if [ -d input ]; then
                        rsync -a --ignore-existing input/ ${APP_DIR}/input/
                    fi

                    cd ${APP_DIR}
                    [ -d venv ] || python3 -m venv venv
                    venv/bin/pip install -q -r requirements.txt
                    venv/bin/playwright install chromium
                    sudo systemctl restart ${SERVICE}
                '''
            }
        }

        stage('Verify') {
            when { branch 'main' }
            steps {
                sh '''
                    sleep 5
                    sudo systemctl is-active ${SERVICE}
                    journalctl -u ${SERVICE} -n 15 --no-pager || true
                '''
            }
        }
    }

    post {
        success { echo 'Deployed successfully' }
        failure { echo 'Pipeline failed, check the stage logs' }
    }
}