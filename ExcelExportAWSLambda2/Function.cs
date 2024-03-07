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
            // クエリ文字列を取得
            string? queryString;
            input.QueryStringParameters.TryGetValue("name", out queryString);

            // ワークシートに追加するテキスト
            string Message = string.IsNullOrEmpty(queryString)
                ? "Hello, World!!"
                : $"Hello, {queryString}!!";

            //Workbook.SetLicenseKey("製品版またはトライアル版のライセンスキーを設定");

            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);

                // S3にアップロード
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
                Body = "ファイルが保存されました。",
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
