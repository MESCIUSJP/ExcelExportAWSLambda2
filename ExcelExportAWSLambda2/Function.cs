using Amazon;
using Amazon.Lambda.APIGatewayEvents;
using Amazon.Lambda.Core;
using Amazon.S3;
using Amazon.S3.Model;
using GrapeCity.Documents.Excel;
using System.Net;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace ExcelExportAWSLambda2;

public class Function
{
    public async Task<APIGatewayProxyResponse> FunctionHandler(APIGatewayProxyRequest input, ILambdaContext
 context)
    {
        APIGatewayProxyResponse response;

        try
        {
            // �N�G����������擾
            string? queryString;
            input.QueryStringParameters.TryGetValue("name", out queryString);

            // ���[�N�V�[�g�ɒǉ�����e�L�X�g
            string Message = string.IsNullOrEmpty(queryString)
                ? "Hello, World!!"
                : $"Hello, {queryString}!!";

            //Workbook.SetLicenseKey("���i�ł܂��̓g���C�A���ł̃��C�Z���X�L�[��ݒ�");

            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);

                // S3�ɃA�b�v���[�h
                AmazonS3Client client = new AmazonS3Client(RegionEndpoint.APNortheast1);
                var request = new PutObjectRequest
                {
                    BucketName = "diodocs-file-export",
                    Key = "Result.xlsx",
                    InputStream = ms
                };

                await client.PutObjectAsync(request);
            }

            response = new APIGatewayProxyResponse
            {
                StatusCode = (int)HttpStatusCode.OK,
                Body = "�t�@�C�����ۑ�����܂����B",
                Headers = new Dictionary<string, string> {
                { "Content-Type", "text/plain; charset=utf-8" }
            }
            };
        }
        catch (Exception e)
        {
            response = new APIGatewayProxyResponse
            {
                StatusCode = (int)HttpStatusCode.InternalServerError,
                Body = e.Message,
                Headers = new Dictionary<string, string> {
                { "Content-Type", "text/plain" }
            }
            };
        }

        return response;
    }
}
