<?php

namespace App\Http\Controllers;

use App\Models\Parcel;
use App\Models\User;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Hash;
use PHPOpenSourceSaver\JWTAuth\Facades\JWTAuth;

class AuthController extends Controller
{

    public function register(Request $request){

        $this->validate($request, [
            'name' => 'required|string|max:100',
            'email' => 'required|string|email|max:255|unique:users',
            'password' => 'required|string|min:6'
        ]);

        // 2. Create the user inside your database
        $user = User::create([
            'name' => $request->input('name'),
            'email' => $request->input('email'),
            'password' => Hash::make($request->input('password')),//+'email'
        ]);


        // 3. Generate the JWT string using the Facade
        $token = JWTAuth::fromUser($user);

        // 4. Calculate the expiration time cleanly using Lumen's config helper
        // Reads your config/jwt.php 'ttl' value (defaults to 60) and converts it to seconds
        $ttlInSeconds = config('jwt.ttl', 60) * 60;

        // 5. Return user data and authorization tokens cleanly
        return $this->success(
            [
                'user'         => $user,
                'access_token' => $token,
                'token_type'   => 'bearer',
                'expires_in'   => $ttlInSeconds
            ],
            'User successfully registered',201// Explicit HTTP Status for account creation
        );


    }

    public function login(Request $request)
    {
        // 1. Validate the inputs
        $this->validate($request, [
            'email' => 'required|string|email|max:255',
            'password' => 'required|string|min:6'
        ]);

        // 2. Locate the user by email using standard Eloquent query syntax
        $user = User::where('email', $request->input('email'))->first();

        // 3. Verify the user exists and the bcrypt password matches
        if (! $user || ! Hash::check($request->input('password'), $user->password)) {
            return response()->json([
                'error' => 'Unauthorized',
                'message' => 'Invalid email or password combination.'
            ], 401);
        }

        // 4. Generate the JWT string directly from the verified user object
        // 🚀 This bypasses Lumen's RequestGuard completely and avoids any crashes!
        $token = JWTAuth::fromUser($user);

        // 5. Return the clean token response


        return $this->success(
            [
                'access_token' => $token,
                'token_type' => 'bearer',
                'expires_in' => config('jwt.ttl', 60) * 60
            ],
            'successfull login',200// Explicit HTTP Status for account creation
        );
    }
    public function me()
    {
        // auth()->user() reads the JWT token from the request header,
        // decrypts the user ID, and fetches the User model from the database.
        return $this->success(auth()->user(),"success",200);
    }

    /**
     * Log the user out (Invalidate the token).
     *
     * Endpoint: POST /api/logout
     */
    public function logout()
    {
        // auth()->logout() immediately puts the current JWT token into a
        // database/cache blacklist so it can never be used again.
        auth()->logout();

        return $this->success(null,
            'Successfully logged out'
        ,200);
    }

    public function five_hundred()
    {
        return ThisFakeFunctionWillCrashPHPNow();

    }


}
