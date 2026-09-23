#!/usr/bin/env bash
# deploy-hosting 워크플로용 키 없는 Google 인증(Workload Identity Federation) 1회 준비.
# Google Cloud Console 의 Cloud Shell(브라우저, 휴대폰 가능)에 통째로 붙여 넣어 실행한다.
# 여러 번 실행해도 이미 있는 항목은 건너뛴다. 서비스 계정 키는 만들지 않는다.
set -u
PROJECT_ID=ai-edu-seco
PROJECT_NUMBER=58934449948
REPO=flying6man27-wq/AI-Edu
SA_NAME=gh-hosting-deploy
SA="$SA_NAME@$PROJECT_ID.iam.gserviceaccount.com"

gcloud config set project "$PROJECT_ID"
gcloud services enable iamcredentials.googleapis.com sts.googleapis.com firebasehosting.googleapis.com

gcloud iam service-accounts describe "$SA" >/dev/null 2>&1 || \
  gcloud iam service-accounts create "$SA_NAME" --display-name="GitHub Actions Hosting deploy"

gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:$SA" --role="roles/firebasehosting.admin" --condition=None >/dev/null

gcloud iam workload-identity-pools describe github --location=global >/dev/null 2>&1 || \
  gcloud iam workload-identity-pools create github --location=global --display-name="GitHub Actions"

gcloud iam workload-identity-pools providers describe ai-edu --location=global \
  --workload-identity-pool=github >/dev/null 2>&1 || \
  gcloud iam workload-identity-pools providers create-oidc ai-edu --location=global \
    --workload-identity-pool=github --display-name="AI-Edu main" \
    --issuer-uri="https://token.actions.githubusercontent.com" \
    --attribute-mapping="google.subject=assertion.sub,attribute.repository=assertion.repository,attribute.ref=assertion.ref" \
    --attribute-condition="assertion.repository=='$REPO' && assertion.ref=='refs/heads/main'"

gcloud iam service-accounts add-iam-policy-binding "$SA" --role="roles/iam.workloadIdentityUser" \
  --member="principalSet://iam.googleapis.com/projects/$PROJECT_NUMBER/locations/global/workloadIdentityPools/github/attribute.repository/$REPO" >/dev/null

echo
echo "완료 — 공급자: projects/$PROJECT_NUMBER/locations/global/workloadIdentityPools/github/providers/ai-edu"
