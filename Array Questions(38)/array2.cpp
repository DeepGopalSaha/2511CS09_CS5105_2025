#include <bits/stdc++.h>
using namespace std;
pair<int,int> brute_second_min_max(vector<int> a){
    sort(a.begin(), a.end());
    int n=a.size();
    int secondSmall=-1, secondLarge=-1;
    for(int i=0;i<n;i++) if(i>0 && a[i]!=a[i-1]){ secondSmall=a[i]; break; }
    for(int i=n-1;i>=0;i--) if(i<n-1 && a[i]!=a[i+1]){ secondLarge=a[i]; break; }
    return {secondSmall, secondLarge};
}
pair<int,int> optimal_second_min_max(vector<int>& a){
    int n=a.size();
    if(n<2) return {-1,-1};
    int min1=INT_MAX, min2=INT_MAX, max1=INT_MIN, max2=INT_MIN;
    for(int x:a){
        if(x<min1){min2=min1; min1=x;}
        else if(x>min1 && x<min2) min2=x;
        if(x>max1){max2=max1; max1=x;}
        else if(x<max1 && x>max2) max2=x;
    }
    if(min2==INT_MAX) min2=-1;
    if(max2==INT_MIN) max2=-1;
    return {min2, max2};
}
int main(){return 0;}