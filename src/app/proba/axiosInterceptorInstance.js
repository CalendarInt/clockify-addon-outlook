import axios from "axios";

const axiosInterceptorInstance = axios.create({});

// Request interceptor
axiosInterceptorInstance.interceptors.request.use(
  async (config) => {
    // Modify the request config here (add headers, authentication tokens)
    let accessToken = JSON.parse(localStorage.getItem("auth"));
    const url = config.url.split("/api/")[1];
    // if (new Date(accessToken?.expiry_date ?? "") < new Date() && url !=='auth') {
    //   let response = await axios.post(
    //     "https://herring-endless-firmly.ngrok-free.app/api/auth/refresh",
    //     {
    //       refreshToken: accessToken.refresh_token,
    //     }
    //   );
    //   let newAuthObject = JSON.stringify(response.data);
    //   localStorage.setItem("auth", newAuthObject);
    //   accessToken = response.data;
    // }

    // if (accessToken) {
    //   config.headers.Authorization = `${accessToken.token_type} ${accessToken.access_token}`;
    // }
    return config;
  },
  (error) => {
    // Handle request errors here
    return Promise.reject(error);
  }
);

// Response interceptor
axiosInterceptorInstance.interceptors.response.use(
  (response) => {
    // Modify the response data here
    return response;
  },
  (error) => {
    // Handle response errors here
    return Promise.reject(error);
  }
);

export default axiosInterceptorInstance;
