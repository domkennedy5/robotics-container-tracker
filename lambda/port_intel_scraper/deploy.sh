#!/bin/bash
# deploy.sh — package and deploy port-intel-scraper Lambda
# Run from lambda/port_intel_scraper/ directory
# Prerequisites: AWS CLI configured, account 844000647671, us-east-1

set -e
FUNCTION_NAME="robotics-port-intel-scraper"
ROLE_NAME="robotics-port-intel-role"
ACCOUNT_ID="844000647671"
REGION="us-east-1"
S3_BUCKET="robotics-container-tracker"

echo "=== Building deployment package ==="
rm -rf package dist
mkdir -p package

pip install -r requirements.txt -t package/ --quiet

cp handler.py package/

cd package
zip -r ../function.zip . -q
cd ..

echo "=== Checking/creating IAM role ==="
ROLE_ARN="arn:aws:iam::${ACCOUNT_ID}:role/${ROLE_NAME}"

aws iam get-role --role-name "$ROLE_NAME" > /dev/null 2>&1 || {
    echo "Creating IAM role..."
    aws iam create-role \
        --role-name "$ROLE_NAME" \
        --assume-role-policy-document '{
          "Version":"2012-10-17",
          "Statement":[{"Effect":"Allow","Principal":{"Service":"lambda.amazonaws.com"},
          "Action":"sts:AssumeRole"}]
        }' > /dev/null
    aws iam attach-role-policy --role-name "$ROLE_NAME" \
        --policy-arn arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole
    aws iam attach-role-policy --role-name "$ROLE_NAME" \
        --policy-arn arn:aws:iam::aws:policy/AmazonS3FullAccess
    echo "Waiting for role to propagate..."
    sleep 15
}

echo "=== Deploying Lambda function ==="
aws lambda get-function --function-name "$FUNCTION_NAME" > /dev/null 2>&1 && {
    echo "Updating existing function..."
    aws lambda update-function-code \
        --function-name "$FUNCTION_NAME" \
        --zip-file fileb://function.zip > /dev/null
} || {
    echo "Creating new function..."
    aws lambda create-function \
        --function-name "$FUNCTION_NAME" \
        --runtime python3.11 \
        --role "$ROLE_ARN" \
        --handler handler.handler \
        --zip-file fileb://function.zip \
        --timeout 300 \
        --memory-size 256 \
        --environment "Variables={S3_BUCKET=${S3_BUCKET},DB_KEY=tracker.db}" \
        --region "$REGION" > /dev/null
}

echo "=== Setting up EventBridge daily schedule (6 AM ET) ==="
RULE_NAME="robotics-port-intel-daily"
aws events put-rule \
    --name "$RULE_NAME" \
    --schedule-expression "cron(0 11 * * ? *)" \
    --state ENABLED \
    --region "$REGION" > /dev/null

LAMBDA_ARN="arn:aws:lambda:${REGION}:${ACCOUNT_ID}:function:${FUNCTION_NAME}"
aws lambda add-permission \
    --function-name "$FUNCTION_NAME" \
    --statement-id "allow-eventbridge-daily" \
    --action "lambda:InvokeFunction" \
    --principal "events.amazonaws.com" \
    --source-arn "arn:aws:events:${REGION}:${ACCOUNT_ID}:rule/${RULE_NAME}" \
    --region "$REGION" 2>/dev/null || true

aws events put-targets \
    --rule "$RULE_NAME" \
    --targets "Id=1,Arn=${LAMBDA_ARN}" \
    --region "$REGION" > /dev/null

echo ""
echo "=== Deployment complete ==="
echo "Function:  $FUNCTION_NAME"
echo "Schedule:  Daily 6:00 AM ET (cron 0 11 * * ? *)"
echo "S3 bucket: $S3_BUCKET / tracker.db"
echo ""
echo "Test invoke:"
echo "  aws lambda invoke --function-name $FUNCTION_NAME --region $REGION out.json && cat out.json"
