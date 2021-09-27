
import io.restassured.path.json.JsonPath;

public class ReUsableMethods {

	
	public static JsonPath rawToJson(String body)
	{
		JsonPath js1 =new JsonPath(body);
		return js1;
	}
	
}