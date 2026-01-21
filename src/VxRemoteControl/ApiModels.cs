using Newtonsoft.Json;

namespace VxRemoteControl
{
    internal sealed class ApiResponse
    {
        [JsonProperty("ok")]
        public bool Ok { get; set; }

        [JsonProperty("data")]
        public object Data { get; set; }

        [JsonProperty("error")]
        public string Error { get; set; }

        public static ApiResponse Success(object data)
        {
            return new ApiResponse { Ok = true, Data = data };
        }

        public static ApiResponse Fail(string error)
        {
            return new ApiResponse { Ok = false, Error = error };
        }
    }
}
